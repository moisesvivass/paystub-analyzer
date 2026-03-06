import logging

def get_logger(name: str) -> logging.Logger:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s — %(levelname)s — %(message)s',
        handlers=[
            logging.FileHandler('process.log'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(name)