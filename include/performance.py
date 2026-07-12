import time
import tracemalloc
import logging

logger = logging.getLogger(__name__)

def perf_tracker(func):
    def wrapper(*args, **kwargs):
        # Start monitoring
        tracemalloc.start()
        start_time = time.perf_counter()

        result = func(*args, **kwargs)

        # Stop timing and capture memory
        end_time = time.perf_counter()
        current, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()

        execution_time = end_time - start_time
        # Convert bytes to MB
        peak_mb = peak / (1024 * 1024)

        logger.info(f"--- {func.__name__} ---")
        logger.info(f"Execution time: {execution_time:.4f} seconds")
        logger.info(f"Peak memory usage: {peak_mb:.4f} MB")

        return result
    return wrapper