"""Base class for declarative components with async resolution support."""

import asyncio
import dataclasses
from collections.abc import Awaitable
from typing import Any, Self


@dataclasses.dataclass
class _ResolveTask:
    field_name: str
    idx: int | None
    awaitable: Awaitable[Any]


def _collect_resolve_tasks(component: "Component") -> list[_ResolveTask]:
    """Collect all async resolution tasks from a component's fields.

    Args:
        component: The component to collect tasks from.

    Returns:
        List of resolution tasks for all async children.
    """
    tasks: list[_ResolveTask] = []
    for field in dataclasses.fields(component):
        value = getattr(component, field.name)
        if value is None:
            continue

        if isinstance(value, Component):
            tasks.append(_ResolveTask(field.name, None, value.resolve()))
        elif asyncio.iscoroutine(value):
            tasks.append(_ResolveTask(field.name, None, value))
        elif isinstance(value, list):
            for idx, item in enumerate(value):
                if isinstance(item, Component):
                    tasks.append(_ResolveTask(field.name, idx, item.resolve()))
                elif asyncio.iscoroutine(item):
                    tasks.append(_ResolveTask(field.name, idx, item))

    return tasks


def _is_value_resolved(value: Any) -> bool:
    """Check if a single value is fully resolved.

    Args:
        value: The value to check.

    Returns:
        True if the value is resolved, False if it's an unresolved Component or coroutine.
    """
    if isinstance(value, Component):
        return value.is_resolved()
    if asyncio.iscoroutine(value):
        return False
    return True


@dataclasses.dataclass
class Component:
    """Base class for all declarative document components.

    Supports both synchronous and asynchronous usage. When any child
    is a coroutine, the component must be awaited to resolve all
    async children concurrently before saving.
    """

    def __await__(self):
        """Make this component awaitable to resolve all async children."""
        return self.resolve().__await__()

    async def resolve(self) -> Self:
        """Recursively resolve all async children concurrently.

        Returns:
            Self with all coroutines replaced by their resolved values.
        """
        tasks = _collect_resolve_tasks(self)

        if tasks:
            results = await asyncio.gather(*(task.awaitable for task in tasks))
            for task, result in zip(tasks, results, strict=True):
                if task.idx is None:
                    setattr(self, task.field_name, result)
                else:
                    getattr(self, task.field_name)[task.idx] = result

        return self

    def is_resolved(self) -> bool:
        """Check if all children are fully resolved (no pending coroutines).

        Returns:
            True if all children are resolved, False otherwise.
        """
        for field in dataclasses.fields(self):
            value = getattr(self, field.name)

            if value is None:
                continue

            if isinstance(value, list):
                if not all(_is_value_resolved(item) for item in value):
                    return False
            elif not _is_value_resolved(value):
                return False

        return True
