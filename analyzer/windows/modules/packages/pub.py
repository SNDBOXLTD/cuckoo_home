# Copyright (C) 2010-2013 Claudio Guarnieri.
# Copyright (C) 2014-2016 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

from lib.common.abstracts import Package


class PUB(Package):
    """Word analysis package."""
    PATHS = [
        ("ProgramFiles", "Microsoft Office", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office10", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office11", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office12", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office14", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office15", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "Office16", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office 15", "root", "office15", "MSPUB.EXE"),
        ("ProgramFiles", "Microsoft Office", "root", "Office16", "MSPUB.EXE"),
    ]

    def start(self, path):
        self._allow_embedded_flash()

        publisher = self.get_path("Microsoft Office Publisher")
        return self.execute(
            publisher, args=["/o", path], mode="office", trigger="file:%s" % path
        )
