# Copyright (C) 2016-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import logging
import socket

from lib.api.process import subprocess_checkcall
from lib.common.abstracts import Auxiliary
from lib.common.exceptions import CuckooPackageError

log = logging.getLogger(__name__)


class DnsFlush(Auxiliary):
    """Flush DNS Cache.

    Flush DNS cache before analysis start so all requests start with a dns query.
    This helps accurate the post processing pcap filtering
    """

    def start(self):
        # verify network
        hostname = "sndbox.com"
        try:
            socket.gethostbyname(hostname)
        except:
            log.exception("Failed to verify network connection.")
        # dns flush
        try:
            subprocess_checkcall(["ipconfig.exe", "/flushdns"])
            log.info("Successfully flushed dns.")
        except Exception as e:
            log.error("Error flushing dns: %s", e)
