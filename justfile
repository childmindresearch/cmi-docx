agentcheck:
    uv run ruff check . --fix
    uv run ty check .
    uv run sergey check .
