try:
    import logging
    import functools
    import time
except:
    raise Exception("Could not load required python libraries. Please run 'pip install -r requirements.txt' then try again.")

#Set logging options.
logger = logging.getLogger("__main__")

def callLogger(func):
    @functools.wraps(func)
    def wrapper_callLogger(*args, **kwargs):
        logger.debug(f'Calling {func.__name__}')
        value = func(*args, **kwargs)
        logger.debug(f'Called {func.__name__}')
        return value
    return wrapper_callLogger

def timer(func):
    @functools.wraps(func)
    def wrapper_timer(*args, **kwargs):
        tic = time.perf_counter()
        value = func(*args, **kwargs)
        toc = time.perf_counter()
        elapsed_time = toc - tic
        logger.debug(f'Elapsed time: {elapsed_time:0.4f} seconds')
        return value
    return wrapper_timer