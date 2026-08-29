"""
logger.py

Application Logger
RPGF Log Card Generator v2.0
"""

import logging
from datetime import datetime

from config import (
    APP_NAME,
    LOG_DIR
)


class Logger:

    def __init__(self):

        log_name = datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S.log"
        )

        log_file = LOG_DIR / log_name

        self.logger = logging.getLogger(APP_NAME)

        self.logger.setLevel(logging.INFO)

        # Remove duplicate handlers
        self.logger.handlers.clear()

        formatter = logging.Formatter(

            "%(asctime)s | %(levelname)-8s | %(message)s",

            "%Y-%m-%d %H:%M:%S"

        )

        file_handler = logging.FileHandler(

            log_file,

            encoding="utf-8"

        )

        file_handler.setFormatter(formatter)

        self.logger.addHandler(file_handler)

    # --------------------------------------------------
    # Information
    # --------------------------------------------------

    def info(self, message):

        self.logger.info(message)

    # --------------------------------------------------
    # Warning
    # --------------------------------------------------

    def warning(self, message):

        self.logger.warning(message)

    # --------------------------------------------------
    # Error
    # --------------------------------------------------

    def error(self, message):

        self.logger.error(message)

    # --------------------------------------------------
    # Critical
    # --------------------------------------------------

    def critical(self, message):

        self.logger.critical(message)

    # --------------------------------------------------
    # Divider
    # --------------------------------------------------

    def line(self):

        self.logger.info(

            "-" * 70

        )

    # --------------------------------------------------
    # Session Header
    # --------------------------------------------------

    def session(self, title):

        self.line()

        self.logger.info(title)

        self.line()