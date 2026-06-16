"""Standalone MCP server entrypoint."""

from excel_tools.mcp_instance import excel_mcp


def register_tools() -> None:
    """Import tool modules so their decorators register with FastMCP."""
    import excel_tools.inspect_excel  # noqa: F401
    import excel_tools.profile_structure  # noqa: F401
    import excel_tools.read_excel  # noqa: F401
    # excel_tools.unmerged_cells is intentionally not imported: it modifies files.


def main() -> None:
    register_tools()
    excel_mcp.run()


if __name__ == "__main__":
    main()
