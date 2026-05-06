import logging
from pathlib import Path
from datetime import datetime


def setup_logger(logs_dir: Path) -> logging.Logger:
    logs_dir.mkdir(exist_ok=True)

    logger = logging.getLogger("robo_stur")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )

    log_file = logs_dir / f"robo_stur_{datetime.now():%Y%m%d_%H%M%S}.log"

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info("Log iniciado em: %s", log_file)
    return logger