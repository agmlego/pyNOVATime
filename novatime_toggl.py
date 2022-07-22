# -*- coding: utf-8 -*-
# SPDX-License-Identifier: FAFOL
# pylint: disable=logging-fstring-interpolation

"""Toggl interface"""

import configparser
import logging
from os.path import isdir, isfile, join

import arrow
import requests
from rich import print
from rich.console import Console, ConsoleOptions, RenderResult
from rich.logging import RichHandler
from rich.progress import Progress
from rich.table import Table

FORMAT = "%(message)s"
logging.basicConfig(
    level="DEBUG",
    format=FORMAT,
    datefmt="[%X]",
    handlers=[RichHandler(rich_tracebacks=True,
                          tracebacks_suppress=[requests])]
)
logging.getLogger("requests").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.WARNING)

class Toggl:

    def __init__(self) -> None:
        self.logger = logging.getLogger('Toggl')
        # set up secrets info
        if not isfile('secrets.ini'):
            raise FileNotFoundError(
                'Please copy secrets.ini.example to secrets.ini and configure per the comments')

        self.logger.debug('Loading secrets file')
        self._secrets = configparser.ConfigParser()
        self._secrets.optionxform = lambda option: option  # return case-sensitive keys
        self._secrets.read('secrets.ini')
        self._api_key = self._secrets['user']['toggl_auth'] 
