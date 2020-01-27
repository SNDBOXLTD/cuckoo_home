# Copyright (C) 2010-2013 Claudio Guarnieri.
# Copyright (C) 2014-2016 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import os
import base64

from lib.common.abstracts import Package


class Dll(Package):
    """DLL analysis package."""

    def _execute_exports(self, path, architecture):
        sndll = os.path.join("bin", "sndll32.exe" if architecture == '32' else "sndll64.exe")

        # Check file extension.
        ext = os.path.splitext(path)[-1].lower()

        # If the file doesn't have the proper .dll extension force it
        # and rename it. This is needed for rundll32 to execute correctly.
        # See ticket #354 for details.
        if ext != ".dll":
            new_path = path + ".dll"
            os.rename(path, new_path)
            path = new_path

        args = ["%s" % path]

        export_settings = self.options.get("dll_export")
        if export_settings:
            # handle execution of a specific export
            decode_settings = base64.b64decode(export_settings)
            args.extend(decode_settings.split())

        return self.execute(sndll, args=args)

    def start(self, path):
        return self._execute_exports(path, '32')
