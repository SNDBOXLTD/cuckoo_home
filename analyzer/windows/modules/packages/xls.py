# Copyright (C) 2012-2013 Claudio Guarnieri.
# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

from lib.common.abstracts import Package
from lib.common.rand import random_string


class XLS(Package):
    """Excel analysis package.
    We hack excel single instance by loading sample from commandline
    """

    PATHS = [
        ("System32", "cmd.exe"),
    ]

    def start(self, path):
        self._allow_embedded_flash()
        # start sample similar to generic package
        cmd_path = self.get_path("cmd.exe")
        # Create random cmd.exe window title.
        rand_title = random_string(4, 16)
        args = ["/c", "start", "/wait", '"%s"' % rand_title, path]
        return self.execute(cmd_path, args=args, trigger="file:%s" % path)
