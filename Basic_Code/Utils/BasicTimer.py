import functools
import time

def timer(func):
    @functools.wraps(func)
    def _wrapper(*args, **kwargs):
        start = time.perf_counter()
        result = func(*args, **kwargs)
        runtime = time.perf_counter() - start
        print(f"{func.__name__} took {runtime:.4f} secs")
        return result
    return _wrapper


if __name__ == "__main__":
    @timer
    def Calc():
        i = 0
        for _ in range(1_000_000):
            i += 1
        return i

    Calc()