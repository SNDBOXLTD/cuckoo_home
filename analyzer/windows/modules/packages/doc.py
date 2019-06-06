# Copyright (C) 2012-2013 Claudio Guarnieri.
# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

from _winreg import HKEY_CURRENT_USER

from lib.common.abstracts import Package

class DOC(Package):
    """Word analysis package."""
    PATHS = [
        ("ProgramFiles", "Microsoft Office", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office10", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office11", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office12", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office14", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office15", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office16", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office 15", "root", "office15", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
        ("ProgramFiles", "Microsoft Office", "WORDVIEW.EXE"),
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
            "Software\\Microsoft\\Office\\14.0\\Word\\Security",
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
            "Software\\Microsoft\\Office\\14.0\\Word\\Security\\FileValidation",
            {
                # Disable file validation check onload
                "EnableOnLoad": 0
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

        word = self.get_path("Microsoft Office Word")
        return self.execute(
            word, args=["/q", path], mode="office", trigger="file:%s" % path
        )
