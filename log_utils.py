import inspect
import logging
import traceback


def call_logger(error):
    caller_function_name = inspect.stack()[1].function
    logging.basicConfig(
        filename="error.log",
        level=logging.ERROR,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filemode="a",
    )
    logger = logging.getLogger(__name__)
    logger.error(f"In {caller_function_name}: {str(error)}")
    logger.error(traceback.format_exc())
