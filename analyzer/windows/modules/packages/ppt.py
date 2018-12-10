# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import os
import logging

from _winreg import HKEY_CURRENT_USER

from lib.common.abstracts import Package

log = logging.getLogger(__name__)

class PPT(Package):
    """PowerPoint analysis package."""
    PATHS = [
        ("ProgramFiles", "Microsoft Office", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office10", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office11", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office12", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office14", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office15", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office16", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office 15", "root", "office15", "POWERPNT.EXE"),
        ("ProgramFiles", "Microsoft Office", "root", "Office16", "POWERPNT.EXE"),
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
            "Software\\Microsoft\\Office\\14.0\\Powerpoint\\Security",
            {
                # Enable VBA macros in Office 2010.
                "VBAWarnings": 1,
                "AccessVBOM": 1,

                # "The file you are trying to open .xyz is in a different
                # format than specified by the file extension. Verify the file
                # is not corrupted and is from trusted source before opening
                # the file. Do you want to open the file now?"
                "ExtensionHardening": 0,
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
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\PowerPoint\\Options",
            {
                # Disable Hardware notification
                "DisableHardwareNotification": 1,
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Control Panel\\Desktop",
            {
                # This value controls how many seconds Windows waits
                # before considering applications unresponsive
                #
                # value in milliseconds
                "HungAppTimeout": "30000",
            },
        ],
    ]

    def start(self, path):
        # We observed multiple files that Powerpoint failed to open
        # directly into slideshow mode (where the exploit occur).
        # Renaming to .pps extention force this situation
        # regarding of prior extention
        if path.endswith(".ppt") or path.endswith(".pptx"):
            os.rename(path, path + ".pps")
            path += ".pps"
            log.info("Submitted file is using ppt/x extension, added .pps")

        self._allow_embedded_flash()

        powerpoint = self.get_path("Microsoft Office PowerPoint")
        return self.execute(
            powerpoint, args=["/S", path], mode="office",
            trigger="file:%s" % path
        )
