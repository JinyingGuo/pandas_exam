from loguru import logger

def log_add(name):
    logger.remove()
    logger.add('../log/' + str(name) + '_log.log',
               level="DEBUG",
               encoding='utf-8')