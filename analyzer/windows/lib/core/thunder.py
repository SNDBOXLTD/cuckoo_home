import os
import platform
import subprocess
import logging
import win32file
import time
import win32process
import win32con
from win32api import FormatMessage
import threading
from ctypes import *
import ctypes
import struct
from lib.common.defines import STARTUPINFO, PROCESS_INFORMATION, PSAPI
from lib.common.defines import CREATE_NEW_CONSOLE, CREATE_SUSPENDED
from lib.common.constants import SHUTDOWN_MUTEX

KERNEL32 = ctypes.windll.kernel32

PACKAGE_TO_PRELOADED_APPS = {
    'XLS': 'excel.exe',
    'DOC': 'winword.exe',
    'PPT': 'powerpnt.exe',
}

# Set logger
log = logging.getLogger(__name__)


def get_error_string(data):
    pass


def get_preloaded_pids():
    """ Fetch list of preloaded pids
    reference: https://code.activestate.com/recipes/305279-getting-process-information-on-windows.

    Returns:
        {dict} -- map of preloaded process name to pid. e.g. {"winword.exe": 312}
    """
    arr = c_ulong * 256
    lpidProcess = arr()
    cb = sizeof(lpidProcess)
    cbNeeded = c_ulong()
    hModule = c_ulong()
    count = c_ulong()
    modname = c_buffer(30)
    PROCESS_QUERY_INFORMATION = 0x0400
    PROCESS_VM_READ = 0x0010

    # Call Enumprocesses to get hold of process id's
    PSAPI.EnumProcesses(byref(lpidProcess), cb, byref(cbNeeded))

    # Number of processes returned
    nReturned = cbNeeded.value/sizeof(c_ulong())
    pidProcess = [i for i in lpidProcess][:nReturned]

    found = dict()

    for pid in pidProcess:
        # Get handle to the process based on PID
        hProcess = KERNEL32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
        if hProcess:
            PSAPI.EnumProcessModules(hProcess, byref(hModule), sizeof(hModule), byref(count))
            PSAPI.GetModuleBaseNameA(hProcess, hModule.value, modname, sizeof(modname))
            process_name = "".join([i for i in modname if i != '\x00'])
            # save process
            if process_name.lower() in PACKAGE_TO_PRELOADED_APPS.values():
                found[process_name.lower()] = pid

            # -- Clean up
            for i in range(modname._length_):
                modname[i] = '\x00'

            KERNEL32.CloseHandle(hProcess)
    return found


class SignatureBuffer(object):
    MONITOR_REASON_FIRST_CREATE = 0
    MONITOR_REASON_CREATE = 1
    MONITOR_REASON_THREAD = 2
    MONITOR_REASON_MEMORY = 3
    MONITOR_REASON_RPC = 4
    MONITOR_REASON_ANY = 0xFFFFFFFF

    PROCESS_RELATIONSHIP_SIG = [
        # ("parent.exe", "child.exe", MONITOR_REASON_MEMORY),  # Injection
        ("winword.exe", "cmd.exe", MONITOR_REASON_ANY),
        ("winword.exe", "powershell.exe", MONITOR_REASON_ANY),
        ("winword.exe", "mshta.exe", MONITOR_REASON_ANY),
        ("winword.exe", "excel.exe", MONITOR_REASON_ANY),
        ("winword.exe", "wscript.exe", MONITOR_REASON_ANY),
        ("winword.exe", "msiexec.exe", MONITOR_REASON_ANY),
        ("winword.exe", "javaw.exe", MONITOR_REASON_ANY),
        ("excel.exe", "certutil.exe", MONITOR_REASON_ANY),
        ("excel.exe", "cmd.exe", MONITOR_REASON_ANY),
        ("excel.exe", "powershell.exe", MONITOR_REASON_ANY),
        ("excel.exe", "wscript.exe", MONITOR_REASON_ANY),
        ("excel.exe", "winword.exe", MONITOR_REASON_ANY),
        ("excel.exe", "msiexec.exe", MONITOR_REASON_ANY),
        ("excel.exe", "mshta.exe", MONITOR_REASON_ANY),
        ("excel.exe", "javaw.exe", MONITOR_REASON_ANY),
        ("WmiPrvSe.exe", "powershell.exe", MONITOR_REASON_ANY),
        ("excel.exe", "WMIC.exe", MONITOR_REASON_ANY),
        ("winword.exe", "WMIC.exe", MONITOR_REASON_ANY)
    ]

    SINGLE_PROCESS_SIG = [
        ("EQNEDT32.EXE", MONITOR_REASON_ANY),  # Equetion editor
        ("explorer.exe", MONITOR_REASON_THREAD),
        ("CaLc.EXE", MONITOR_REASON_ANY),
        ("powershell.exe", MONITOR_REASON_CREATE),
        ("cmd.exe", MONITOR_REASON_CREATE),
        ("wscript.exe", MONITOR_REASON_CREATE)
    ]

    TYPE_ONE_PROCESS = 1
    TYPE_RELATIONSHIP = 2

    def __init__(self):
        self.buffer = self.get_sig_buff()

    def get_sig_buff(self):
        buffer = ""
        # One process

        # Single process
        for proc_name, reason in self.SINGLE_PROCESS_SIG:
            proc_name = proc_name.encode("utf-16-le")  # Unicode WCHAR

            s = ""
            s += struct.pack("I", self.TYPE_ONE_PROCESS)  # Type
            s += struct.pack("I", reason)
            s += struct.pack("64s", proc_name)

            # Set length as first parameter
            s = struct.pack("I", len(s) + 4) + s

            # Add to buffer
            buffer += s

        # Relationship
        for parent_proc_name, child_proc_name, reason in self.PROCESS_RELATIONSHIP_SIG:
            parent_proc_name = parent_proc_name.encode("utf-16-le")  # Unicode WCHAR
            child_proc_name = child_proc_name.encode("utf-16-le")  # Unicode WCHAR

            s = ""
            s += struct.pack("I", self.TYPE_RELATIONSHIP)
            s += struct.pack("I", reason)
            s += struct.pack("64s", parent_proc_name)
            s += struct.pack("64s", child_proc_name)

            # Set length as first parameter
            s = struct.pack("I", len(s) + 4) + s

            # Add to buffer
            buffer += s

        # print buffer
        # log.debug(buffer)
        # log.debug(buffer.encode("hex"))
        return buffer


class Thunder(object):
    def __init__(self, pipe_name, forwarder_pipe_name, dispatcher_pipe_name, destination, package, configuration):
        self.is_x64 = platform.machine().endswith("64")
        self._driver_communication_device = 0
        self.ip, self.port = destination

        # Ioctls
        self._ioctl_monitor = 0x222408
        self._ioctl_configuration = 0x22240C
        self._ioctl_communication_new_pipe_name = 0x222410
        self._ioctl_stop_monitoring = 0x22241C
        self._ioctl_thunder_sig_process = 0x222428

        # Order is crucial, same in the driver it self
        self._configuration = configuration
        self._configuration_order = ["SSDT", "TIME", "REGISTRY", "FILES",
                                     "EXTRA", "LOGGING", "AGGRESSIVE", "RPC", "ULTRAFAST", "LIGHT", "MEMDUMP"]

        # General configurations
        self._driver_pipe_name = "\\\\.\\Thunder"
        self._driver_log_pipe_name = pipe_name  # Kernel

        # Cuckoo is reading from this pipe to forward the host machine
        self._forwarder_log_pipe_name = forwarder_pipe_name  # Forwarder
        self._dispatcher_log_pipe_name = dispatcher_pipe_name  # Dispatcher

        # Binary configuration, exactly as in the binary directory
        self._installer_dll_name = "WdfCoinstaller01009.dll"
        self._installer_exe_name = "Strike.exe"
        self._driver_name = "Thunder.sys"
        self._information_file = "minimal.inf"
        self._log_dispatcher_name = "log_dispatcher.pyw"

        # holds pid of preloaded office apps
        self.package = package
        log.debug("thunder using package: %s", self.package)
        self.preloaded_pids = get_preloaded_pids()

    def _check_preloaded_pid(self, pid):
        """check preloaded pid. 
        Arguments:
            pid {int} -- pid of injected process

        Returns:
            {int} -- pid of preloaded app or original pid
        """
        preloaded_app = PACKAGE_TO_PRELOADED_APPS.get(self.package)
        return self.preloaded_pids.get(preloaded_app, pid)

    def _create_device(self):
        # return KERNEL32.CreateFileA(self._driver_pipe, GENERIC_READ | GENERIC_WRITE, 0, None, OPEN_EXISTING, 0, None)
        return win32file.CreateFile(self._driver_pipe_name, win32file.GENERIC_READ | win32file.GENERIC_WRITE, 0, None,
                                    win32file.OPEN_EXISTING, 0, None)

    def _send_ioctl(self, device, ioctl, msg):
        to_send = msg
        if long == type(msg) or int == type(msg):
            # to_send = ("0" + hex(msg)[2:]).replace("L", "").decode("hex")
            to_send = ("%08x" % (msg)).decode("hex")
            length = len(to_send)
        else:
            length = len(str(msg))

        print "Sending: [%s] of length: [%d]" % (to_send, length)
        print type(to_send)
        # return KERNEL32.DeviceIoControl(device, ioctl, to_send, length, None) # Not working with kernel32 like that
        return win32file.DeviceIoControl(device, ioctl, to_send, length, None)

    def _place_driver(self):
        """
        Places the driver for the current architecture in "bin/Thunder.sys"
        Remove other driver version

        The installed driver expects to be called "Thunder.sys" on disk, therefore trying to install a driver with a
        different name directly will not work.
        :return: None
        """
        matching_arch_driver = "Thunder64.sys" if self.is_x64 else "Thunder32.sys"
        to_remove_driver = "Thunder64.sys" if not self.is_x64 else "Thunder32.sys"
        os.rename(os.path.join("bin", matching_arch_driver), os.path.join("bin", self._driver_name))
        os.remove(os.path.join("bin", to_remove_driver))

    def check_components(self):
        installing_components = [
            self._installer_dll_name,
            self._installer_exe_name,
            self._driver_name,
            self._information_file,
            self._log_dispatcher_name
        ]

        # Sanity
        for component in installing_components:
            comp_path = os.path.abspath(os.path.join("bin", component))
            if not os.path.exists(comp_path):
                log.warning("Driver component not found: [%s]" % (comp_path))
                return False
        return True

    def install(self):
        log.info("Thunder - Installation initialized")

        self._place_driver()

        # Sanity
        if not self.check_components():
            return False

        # Initialize logger command
        args_logs = [
            "cmd.exe",
            "/c",
            "start",
            os.path.abspath(os.path.join("bin", self._log_dispatcher_name)),
            str(self._driver_log_pipe_name),
            str(self._forwarder_log_pipe_name),
            str(self._dispatcher_log_pipe_name),
            self.ip,
            str(self.port)
        ]

        # RUNDLL32.EXE SETUPAPI.DLL,InstallHinfSection DefaultInstall 132 C:\temp\minimal.inf
        args_inf_installer = [
            "C:\windows\system32\RUNDLL32.EXE",
            "SETUPAPI.DLL,InstallHinfSection",
            "DefaultInstall",
            "132",
            os.path.abspath(os.path.join("bin", self._information_file)),
        ]

        # Initialize installer command
        # Strike.exe 1 C:\temp\WdfCoinstaller01009.dll C:\temp\minimal.inf C:\temp\thunder.sys
        args_installer = [
            os.path.abspath(os.path.join("bin", self._installer_exe_name)),
            "1",
            os.path.abspath(os.path.join("bin", self._installer_dll_name)),
            os.path.abspath(os.path.join("bin", self._information_file)),
            os.path.abspath(os.path.join("bin", self._driver_name)),
        ]

        # Execute command
        log.info("Execution args: [%s][%s]" % (args_logs, args_installer))
        try:
            subprocess.check_call(args_logs, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            if self.is_x64:
                KERNEL32.Wow64DisableWow64FsRedirection(0)
            subprocess.check_call(args_inf_installer, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                  stderr=subprocess.PIPE)

            subprocess.check_call(args_installer, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError:
            log.error("Failed [CalledProcessError] installing driver with command args: [%s][%s][%s]" % (
                args_logs, args_inf_installer, args_installer))
            return False
        log.info("Driver installed successfully")

        return self.initialize()

    @staticmethod
    def errcheck(result, func, args):
        """
        Kernel32 hook, if an error occurred, we rise an exception explaining
        what has happened, the exception will later be caught be the Sndbox cuckoo module.

        returns the argument tuple it receives unchanged,
        ctypes continues the normal processing it does on the output parameters.

        :param result: result is what the foreign function returns
        :param func: func is the foreign function object itself
        :param args: a tuple containing the parameters originally passed to the function call
        :return: function arguments
        """
        if not result:
            error_code = KERNEL32.GetLastError()
            error_message = "error: ({}, '{}', '{}')".format(error_code,
                                                             func.__name__,  # for example, CreateProcessW
                                                             FormatMessage(error_code).strip())
            log.error(error_message)
            raise Exception(error_message)
        return args

    def create_monitored_process(self, process_path, args=""):
        """
        Creating a new process to monitor by the driver.
        The process will be created in suspended mode, then after the PID will be
        sent to the driver for further monitoring. When the driver is done
        initializing, the process can be resumed.

        Parameters:
            process_path (string), the path of the process to create, unicode
                string.
            args (unicode), additional arguments to provide to the process

        Return value:
            list of process parameters: (h_process, h_thread, pid, tid)
        """

        # Initializations
        startup_info = STARTUPINFO()
        startup_info.cb = sizeof(startup_info)
        startup_info.dwFlags = win32process.STARTF_USESHOWWINDOW
        startup_info.wShowWindow = win32con.SW_NORMAL
        KERNEL32.CreateProcessW.errcheck = self.errcheck  # assign a callable

        process_info = PROCESS_INFORMATION()

        creation_flags = CREATE_NEW_CONSOLE | CREATE_SUSPENDED
        # Create new process (UNICODE)
        success = KERNEL32.CreateProcessW(None,
                                          "%s %s" % (process_path, args),
                                          None,
                                          None,
                                          None,
                                          creation_flags,
                                          None,
                                          None,  # Starting path?
                                          ctypes.byref(startup_info),
                                          ctypes.byref(process_info))

        # Data
        h_process, h_thread, pid, tid = process_info.hProcess, process_info.hThread, process_info.dwProcessId, process_info.dwThreadId,
        pid = self._check_preloaded_pid(pid)

        # Hack to monitor first pid - this process
        try:
            self._send_ioctl(self._driver_communication_device, self._ioctl_monitor, str(pid))
            pass
        except Exception, e:
            error_code = KERNEL32.GetLastError()
            log.error("Failed monitoring, GLE: [%d]-[%s]", error_code, get_error_string(error_code))
            log.error(str(e))
            return (False, h_process, h_thread, pid, tid)

        log.info("Process successfully monitored, Now resuming PID[%d] TID[%d].", pid, tid)

        # Resume suspended thread
        if -1 != KERNEL32.ResumeThread(h_thread):
            log.info("Successfully resumed.")
        else:
            log.error("Resume failed.")

        return (True, h_process, h_thread, pid, tid)

    def initialize(self):
        # Create driver device
        while 1:
            try:
                self._driver_communication_device = win32file.CreateFile(self._driver_pipe_name,
                                                                         win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                                                                         0, None,
                                                                         win32file.OPEN_EXISTING, 0, None)
            except:
                log.info("Failed creating communication device: [%d]", KERNEL32.GetLastError())
                time.sleep(1)
                continue
            break
        # Send new pipe for logging communication creation request
        try:
            self._send_ioctl(self._driver_communication_device, self._ioctl_communication_new_pipe_name,
                             self._driver_log_pipe_name.split("\\")[-1] + "\x00")
        except Exception, e:
            error_code = KERNEL32.GetLastError()
            log.error("Failed create_pipe, GLE: [%d]-[%s], dev: [%s]", error_code,
                      get_error_string(error_code), self._driver_log_pipe_name)
            log.error(str(e))
            raw_input()
            return False

        log.info("New pipename initialized: [%s]", self._driver_log_pipe_name)
        return self.initialize_ultrafast_signatures()

    def initialize_ultrafast_signatures(self):
        if not self._configuration.get("ultrafast", False):
            return True

        try:
            # signatures bugger
            signatures_buffer = SignatureBuffer().buffer
            self._send_ioctl(self._driver_communication_device, self._ioctl_thunder_sig_process, signatures_buffer)
        except Exception, e:
            error_code = KERNEL32.GetLastError()
            log.error("Failed initialize_signatures, GLE: [%d]-[%s], dev: [%s]", error_code,
                      get_error_string(error_code), self._driver_log_pipe_name)
            log.error(str(e))
            raw_input()
            return False

        log.info("initialize_signatures initialized: [%s]", self._driver_log_pipe_name)
        return True

    def monitor(self):
        # Initialize device
        binary_conf = ""

        if 0 == self._driver_communication_device:
            log.error("Bad driver pipe device")
            return False

        try:
            # Parse configurations
            binary_conf = self.parse_configuration()
            log.info("Driver configuration is: [0x%08X]" % binary_conf)

            # Send configuration
            self._send_ioctl(self._driver_communication_device, self._ioctl_configuration, binary_conf)

            # Hack to monitor first pid - this process
            # self._send_ioctl(self._driver_communication_device, self._ioctl_monitor, str(os.getpid()))
        except Exception, e:
            error_code = KERNEL32.GetLastError()
            log.error("Failed monitoring, GLE: [%d]-[%s], dev: [0x%08X], conf: [%s]", error_code,
                      get_error_string(error_code), self._driver_communication_device, binary_conf)
            log.error(str(e))
            return False

        log.info("Driver monitor initialized")
        return True

    def parse_configuration(self):
        number = ""
        log.info("Configuration dict: [%s]", str(self._configuration))
        for conf_title in self._configuration_order:
            val = self._configuration.get(
                conf_title.lower(), False) or self._configuration.get(conf_title.upper(), False)
            log.info("Driver Configuration [%s] : [%s]", conf_title, str(val))
            if val:
                number = "1" + number
            else:
                number = "0" + number

        return int(number, 2)

    def thread_wait_finish(self):
        # Wait for shutdown mutex to be created
        while True:
            time.sleep(1)
            # Create the shutdown mutex.
            mutex_handle = KERNEL32.OpenMutexA(0x00100000, False, SHUTDOWN_MUTEX)

            # If shutdown mutex is found, exit loop
            if 0 != mutex_handle:
                KERNEL32.CloseHandle(mutex_handle)
                break

        # Stop monitoring
        self._send_ioctl(self._driver_communication_device, self._ioctl_stop_monitoring, "dummymessage")

    def wait_finish(self):
        t = threading.Thread(target=self.thread_wait_finish)
        t.start()

        return True
