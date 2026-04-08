"""Base class for declarative components with async resolution support."""

import asyncio
import dataclasses
from collections.abc import Awaitable, Callable, Generator
from typing import Any, Self


@dataclasses.dataclass
class _ResolveTask:
    field_name: str
    idx: int | None
    awaitable: Awaitable[Any]


def _materialize_lazy_fields(component: "Component") -> None:
    """Materialize any callable (lazy) field values on a component.

    Fields whose value is callable (but not a Component, coroutine, or future)
    are replaced with the result of calling them. This allows children to be
    passed as ``lambda: [...]`` so their construction is deferred until
    resolution time.

    Args:
        component: The component whose fields should be materialized.
    """
    for field in dataclasses.fields(component):
        value = getattr(component, field.name)
        if (
            callable(value)
            and not isinstance(value, Component)
            and not asyncio.iscoroutine(value)
            and not asyncio.isfuture(value)
            and field.name != "condition"
        ):
            setattr(component, field.name, value())


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
        elif asyncio.iscoroutine(value) or asyncio.isfuture(value):
            tasks.append(_ResolveTask(field.name, None, value))
        elif isinstance(value, (list, tuple, Generator)):
            for idx, item in enumerate(value):
                if isinstance(item, Component):
                    tasks.append(_ResolveTask(field.name, idx, item.resolve()))
                elif asyncio.iscoroutine(item) or asyncio.isfuture(item):
                    tasks.append(_ResolveTask(field.name, idx, item))
    return tasks


@dataclasses.dataclass
class Component:
    """Base class for all declarative document components.

    All components are async. Use `await Document(...)` to resolve all
    async children concurrently.

    Attributes:
        condition: If Callable resolves to False, will not render the component.
    """

    condition: Callable[[], bool] = dataclasses.field(
        default=lambda: True, kw_only=True
    )

    def __await__(self) -> Generator[None, None, Self]:
        """Convenience method for awaiting a component."""
        return self.resolve().__await__()

    async def resolve(self) -> Self:
        """Recursively resolve all async children concurrently.

        Returns:
            Self with all coroutines replaced by their resolved values
                and callables materialized.
        """
        if not self.condition():
            return self

        _materialize_lazy_fields(self)

        tasks = _collect_resolve_tasks(self)

        if tasks:
            results = await asyncio.gather(*(task.awaitable for task in tasks))
            for task, result in zip(tasks, results, strict=True):
                if task.idx is None:
                    setattr(self, task.field_name, result)
                else:
                    getattr(self, task.field_name)[task.idx] = result

        return self
