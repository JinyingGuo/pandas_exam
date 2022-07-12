from loguru import logger
logger.add('../log/my_log.log', encoding='utf-8', level="INFO")