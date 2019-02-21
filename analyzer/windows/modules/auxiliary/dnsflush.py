# Copyright (C) 2016-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import logging

from lib.api.process import subprocess_checkcall
from lib.common.abstracts import Auxiliary
from lib.common.exceptions import CuckooError

log = logging.getLogger(__name__)


class DnsFlush(Auxiliary):
    """Flush DNS Cache."""

    def start(self):
        try:
            subprocess_checkcall(["ipconfig.exe", "/flushdns"])
            log.info("Successfully flushed dns.")
        except CuckooError as e:
            log.error("Error flushing dns: %s", e)
