# Copyright (C) 2012-2013 Claudio Guarnieri.
# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

from _winreg import HKEY_CURRENT_USER

from lib.common.abstracts import Package
from lib.common.rand import random_string

class XLS(Package):
    """Excel analysis package.
    We hack excel single instance by loading sample from commandline
    """

    PATHS = [
        ("System32", "cmd.exe"),
    ]

    REGKEYS = [
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Common\\General",
            {
                # "Welcome to the 2010 Microsoft Office system"
                "ShownOptIn": 1,
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Excel\\Security",
            {
                # Enable VBA macros in Office 2010.
                "VBAWarnings": 1,
                "AccessVBOM": 1,

                # "The file you are trying to open .xyz is in a different
                # format than specified by the file extension. Verify the file
                # is not corrupted and is from trusted source before opening
                # the file. Do you want to open the file now?"
                "ExtensionHardening": 0,

                # "Data connection has been blocked"
                "DataConnectionWarnings": 0,
                "WorkbookLinkWarnings": 0,
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\Common\\Security",
            {
                # Enable all ActiveX controls without restrictions & prompting.
                "DisableAllActiveX": 0,
                "UFIControls": 1,
            },
        ],
    ]

    def start(self, path):
        self._allow_embedded_flash()
        # start sample similar to generic package
        cmd_path = self.get_path("cmd.exe")
         # Create random cmd.exe window title.
        rand_title = random_string(4, 16)
        args = ["/c", "start", "/wait", '"%s"' % rand_title, path]
        return self.execute(cmd_path, args=args, trigger="file:%s" % path)
