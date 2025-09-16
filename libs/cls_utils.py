from threading import RLock
from typing import TypeVar, Type, Any, Dict, Tuple, ClassVar, cast

T = TypeVar('T')


class Singleton(type):
    _instances: ClassVar[Dict[Type[Any], Any]] = {}
    _lock: ClassVar[RLock] = RLock()

    def __call__(cls: Type[T], *args: Any, **kwargs: Any) -> T:
        singleton_cls = cast(Singleton, cls.__class__)
        with singleton_cls._lock:
            if cls not in singleton_cls._instances:
                # Build the first instance of the class
                instance = super(Singleton, singleton_cls).__call__(cls, *args, **kwargs)
                singleton_cls._instances[cls] = instance
            else:
                # An instance of the class already exists
                instance = singleton_cls._instances[cls]
                # Here we are going to call the __init__ and maybe reinitialize
                if getattr(cls, '__allow_reinitialization', False):
                    # If the class allows reinitialization, then do it
                    instance.__init__(*args, **kwargs)
        return cast(T, instance)