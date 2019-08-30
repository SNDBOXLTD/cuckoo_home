# Copyright (C) 2016-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import logging


from _winreg import CreateKey, SetValueEx, CloseKey, REG_DWORD, REG_SZ, HKEY_CURRENT_USER

from lib.common.abstracts import Auxiliary
from lib.common.abstracts import Package
from lib.common.exceptions import CuckooPackageError

log = logging.getLogger(__name__)


class OfficeRegKeys(Auxiliary):
    """Handle RegKeys fix for all office packages.

    This this handle case of office package being open as a child process without proper preparations
    e.g. Word -> Excel with macros
    """

    # Word
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

                # Disable Data Execution Prevention
                "EnableDEP": 0
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Word\\Security\\ProtectedView",
            {
                # Disable Protected View
                "DisableAttachementsInPV": 1,
                "DisableAttachmentsInPV": 1,
                "DisableInternetFilesInPV": 1,
                "DisableUnsafeLocationsInPV": 1
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Word\\Security\\FileBlock",
            {
                # turn off file block for legacy types
                "OpenInProtectedView": 2,
                "Word2000Files": 0,
                "Word60Files": 0,
                "Word95Files": 0,
                "Word2Files": 0,
                "WordXPFiles": 0
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
                
                # Publisher Automation Security Level 
                "automationsecuritypublisher": 1,
            },
        ],
    ]
    # Excel
    REGKEYS += [
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

                # Disable Data Execution Prevention
                "EnableDEP": 0
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Excel\\Security\\FileValidation",
            {
                # Disable file validation check onload
                "EnableOnLoad": 0
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Excel\\Security\\ProtectedView",
            {
                # Disable Protected View
                "DisableAttachementsInPV": 1,
                "DisableAttachmentsInPV": 1,
                "DisableInternetFilesInPV": 1,
                "DisableUnsafeLocationsInPV": 1
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Excel\\Security\\FileBlock",
            {
                # turn off file block for legacy types
                "OpenInProtectedView": 2,
                "xl97workbooksandtemplates": 0,
                "xl97addins": 0,
                "xl9597workbooksandtemplates": 0,
                "xl95workbooks": 0,
                "xl4workbooks": 0,
                "xl4worksheets": 0,
                "xl4macros": 0,
                "xl3worksheets": 0,
                "xl3macros": 0,
                "xl2worksheets": 0,
                "xl2macros": 0
            },
        ],
    ]
    # PUBLISHER
    REGKEYS += [
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Publisher",
            {
                "promptforbadfiles": 1,
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Publisher\\Security",
            {
                # Enable VBA macros in Office 2010.
                "VBAWarnings": 1,
                "AccessVBOM": 1,

                # "The file you are trying to open .xyz is in a different
                # format than specified by the file extension. Verify the file
                # is not corrupted and is from trusted source before opening
                # the file. Do you want to open the file now?"
                "ExtensionHardening": 0,

                # Disable Data Execution Prevention
                "EnableDEP": 0,
            },
        ],
        [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Publisher\\Security\\FileValidation",
            {
                # Disable file validation check onload
                "EnableOnLoad": 0
            },
        ],
                [
            HKEY_CURRENT_USER,
            "Software\\Microsoft\\Office\\14.0\\Word\\Security\\ProtectedView",
            {
                # Disable Protected View
                "DisableAttachmentsInPV": 1,
                "DisableInternetFilesInPV": 1,
                "DisableUnsafeLocationsInPV": 1
            },
        ],
    ]
    # PPT
    REGKEYS += [
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
            "Software\\Microsoft\\Office\\14.0\\Powerpoint\\Security\\FileValidation",
            {
                # Disable file validation check onload
                "EnableOnLoad": 0
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

    def start(self):
        p = Package()
        p.init_regkeys(self.REGKEYS)
        log.info("Successfully installed Office registry keys.")
