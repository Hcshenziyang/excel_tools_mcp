"""Standalone MCP server entrypoint."""

from excel_tools.runtime.mcp import excel_mcp


def register_tools() -> None:
    """Import tool modules so their decorators register with FastMCP."""
    import excel_tools.tools.inspect_excel  # noqa: F401
    import excel_tools.tools.profile_structure  # noqa: F401
    import excel_tools.tools.read_excel  # noqa: F401
    # excel_tools.tools.unmerged_cells is intentionally not imported: it modifies files.


def main() -> None:
    register_tools()
    excel_mcp.run()


if __name__ == "__main__":
    main()
