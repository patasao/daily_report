import logging
from datetime import datetime

def get_logger(name: str, log_dir) -> logging.Logger:
    logger = logging.getLogger(name)

    if not logger.handlers:
        log_dir.mkdir(parents=True, exist_ok=True)

        log_file = log_dir / f'{name}_log_at_{datetime.now().strftime("%d%m%Y%H%M%S")}.log'

        logger.setLevel(logging.INFO)

        handler = logging.FileHandler(log_file, delay=True)
        handler.setFormatter(logging.Formatter(
            fmt='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        logger.addHandler(handler)
        logger.propagate = False

    return logger
