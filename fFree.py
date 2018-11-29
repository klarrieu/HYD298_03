import logging, os


def logging_begin(*logfile_name):
    try:
        logfile_name = logfile_name[0]
    except ValueError:
        logfile_name = "logfile.log"

    logger = logging.getLogger(logfile_name)
    logger.setLevel(logging.INFO)

    # create console handler and set level to info
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter("%(asctime)s - %(message)s")
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    f_handler = logging.FileHandler(os.path.join(os.getcwd(), logfile_name), "w", encoding=None, delay="true")
    f_handler.setLevel(logging.INFO)
    f_formatter = logging.Formatter("%(asctime)s - %(message)s")
    f_handler.setFormatter(f_formatter)
    logger.addHandler(f_handler)

    return logger


def logging_end(logger):
    for handler in logger.handlers:
        handler.close()
        logger.removeHandler(handler)