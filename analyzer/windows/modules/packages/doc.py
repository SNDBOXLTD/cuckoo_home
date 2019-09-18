# Copyright (C) 2012-2013 Claudio Guarnieri.
# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import os
import logging

from lib.common.abstracts import Package
from lib.common.rand import random_string

log = logging.getLogger(__name__)


class DOC(Package):
    """Word analysis package.

    In order to save loading time, we use preloaded winword.exe.
    To prevent winword to open a new process, we launch the file from commandline.
    """

    PATHS = [
        ("System32", "cmd.exe"),
    ]

    def start(self, path):

        # rename dotm to doc, so no new instance is created
        if path.endswith(".dotm"):
            os.rename(path, path + ".doc")
            path += ".doc"
            log.info("Submitted file is using dotm extension, added .doc")

        self._allow_embedded_flash()
        # start sample similar to generic package
        # TODO: Use Generic package code

        cmd_path = self.get_path("cmd.exe")

        # Create random cmd.exe window title.
        rand_title = random_string(4, 16)
        args = ["/c", "start", "/wait", '"%s"' % rand_title, path]
        return self.execute(cmd_path, args=args, trigger="file:%s" % path)
