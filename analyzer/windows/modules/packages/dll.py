# Copyright (C) 2010-2013 Claudio Guarnieri.
# Copyright (C) 2014-2016 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import os
import shlex
import shutil

from lib.common.abstracts import Package

class Dll(Package):
    """DLL analysis package."""
    PATHS = [
        ("bin", "sndll32.exe"),
    ]

    def start(self, path):
        sndll = self.get_path("sndll32.exe")

        # Check file extension.
        ext = os.path.splitext(path)[-1].lower()

        # If the file doesn't have the proper .dll extension force it
        # and rename it. This is needed for rundll32 to execute correctly.
        # See ticket #354 for details.
        if ext != ".dll":
            new_path = path + ".dll"
            os.rename(path, new_path)
            path = new_path

        args = ["%s" % (path)]

        return self.execute(sndll, args=args)
