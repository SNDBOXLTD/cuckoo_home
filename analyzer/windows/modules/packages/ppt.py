# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import os
import logging


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
