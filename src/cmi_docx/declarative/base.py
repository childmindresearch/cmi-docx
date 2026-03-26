"""Base class for declarative components with async resolution support."""

import asyncio
import dataclasses
from typing import TYPE_CHECKING, Any, Self

if TYPE_CHECKING:
    from collections.abc import Awaitable


class Component:
    """Base class for all declarative document components.

    Supports both synchronous and asynchronous usage. When any child
    is a coroutine, the component must be awaited to resolve all
    async children concurrently before saving.
    """

    def __await__(self) -> Any:
        """Make this component awaitable to resolve all async children."""
        return self._resolve().__await__()

    async def _resolve(self) -> Self:
        """Recursively resolve all async children concurrently.

        Returns:
            Self with all coroutines replaced by their resolved values.
        """
        if not dataclasses.is_dataclass(self):
            return self

        tasks: list[tuple[str, int | None, Awaitable[Any]]] = []

        for field in dataclasses.fields(self):
            value = getattr(self, field.name)

            if value is None:
                continue

            if isinstance(value, Component):
                tasks.append((field.name, None, value._resolve()))
            elif asyncio.iscoroutine(value):
                tasks.append((field.name, None, value))
            elif isinstance(value, list):
                for idx, item in enumerate(value):
                    if isinstance(item, Component):
                        tasks.append((field.name, idx, item._resolve()))
                    elif asyncio.iscoroutine(item):
                        tasks.append((field.name, idx, item))

        if tasks:
            results = await asyncio.gather(*(task[2] for task in tasks))

            for (field_name, idx, _), result in zip(tasks, results, strict=True):
                if idx is None:
                    setattr(self, field_name, result)
                else:
                    getattr(self, field_name)[idx] = result

        return self

    def _is_resolved(self) -> bool:
        """Check if all children are fully resolved (no pending coroutines).

        Returns:
            True if all children are resolved, False otherwise.
        """
        if not dataclasses.is_dataclass(self):
            return True

        for field in dataclasses.fields(self):
            value = getattr(self, field.name)

            if value is None:
                continue

            if isinstance(value, Component) and not value._is_resolved():
                return False

            if asyncio.iscoroutine(value):
                return False

            if isinstance(value, list):
                for item in value:
                    if isinstance(item, Component) and not item._is_resolved():
                        return False
                    if asyncio.iscoroutine(item):
                        return False

        return True
