# Excel Tools MCP

Excel Tools MCP is a standalone Model Context Protocol server for reading and inspecting local
Excel files. It is designed for AI assistants that need a compact, structured view of spreadsheets
without loading an entire workbook into context.

The project supports two distribution paths:

- npm / npx: a Node.js launcher starts the Python MCP server.
- PyPI / uvx / pip: run the Python MCP server directly.

Only local file paths are supported.

## Version Status

### Published v0.1.0

The first published release supports `.xlsx` files through `openpyxl`.

Available tools:

- `excel_inspect`: inspect workbook metadata and sheet dimensions.
- `excel_read_range`: read a rectangular cell range.
- `excel_profile_structure`: summarize row structure patterns.

Known boundary:

- `.xls` is not supported in v0.1.0.

### Published v0.1.1

New and changed behavior:

- Read-only tools support `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, and `.ods`.
- `.xlsx` and `.xlsm` are read with `openpyxl`.
- `.xls`, `.xlsb`, and `.ods` are read with `python-calamine`.
- `excel_read_range_normalized` reads a range and virtually fills merged cells from their anchor
  values without modifying the file.

Available tools in v0.1.1:

- `excel_inspect`: inspect workbook metadata and sheet dimensions.
- `excel_read_range_normalized`: read a rectangular range, analyze merged cells, and optionally
  return anchor-filled data.
- `excel_profile_structure`: summarize row structure patterns and merged-cell structure.

## Run With npx

This is the easiest route for MCP clients that already support Node-based server launch commands.

```json
{
  "mcpServers": {
    "excel-tools-mcp": {
      "type": "stdio",
      "command": "npx",
      "args": ["--yes", "@wasziyang/excel-tools-mcp"]
    }
  }
}
```

Requirements:

- Node.js 20+
- Python 3.10+

The npm package is a launcher. It creates or reuses a cached Python environment, installs the Python
MCP server, and starts it over stdio.

Terminal test:

```bash
npx --yes @wasziyang/excel-tools-mcp
```

The command may appear to do nothing. That is normal for an MCP stdio server: it waits for the MCP
host to send JSON-RPC messages over stdin.

### Custom Python Path

By default, the launcher searches for `python3`, then `python`.

Set `EXCEL_TOOLS_MCP_PYTHON` only when Python is installed somewhere unusual:

```json
{
  "mcpServers": {
    "excel-tools-mcp": {
      "type": "stdio",
      "command": "npx",
      "args": ["--yes", "@wasziyang/excel-tools-mcp"],
      "env": {
        "EXCEL_TOOLS_MCP_PYTHON": "/absolute/path/to/python"
      }
    }
  }
}
```

## Run With PyPI / uvx

Users with `uv` can run the PyPI package directly:

```json
{
  "mcpServers": {
    "excel-tools-mcp": {
      "type": "stdio",
      "command": "uvx",
      "args": [
        "--from",
        "excel-tools-mcp",
        "excel-tools-mcp"
      ]
    }
  }
}
```

Terminal test:

```bash
uvx --from excel-tools-mcp excel-tools-mcp
```

You can also install it with pip:

```bash
pip install excel-tools-mcp
excel-tools-mcp
```

Runtime difference:

```text
npx -> npm package -> Node launcher -> Python MCP server
uvx -> PyPI package -> Python MCP server
pip -> PyPI package -> Python MCP server
```

## Local Development

Install from the local checkout:

```bash
pip install -e .
excel-tools-mcp
```

Or run the module directly:

```bash
python3 -m excel_tools.server
```

For a local npx-style test:

```bash
npm start
```

## Tool Arguments Example

```json
{
  "file_path": "/absolute/path/to/report.xlsx",
  "sheet": "Sheet1",
  "start_cell": "A1",
  "end_cell": "D20"
}
```

For Windows paths:

```json
{
  "file_path": "C:\\Users\\Alice\\Documents\\report.xlsx",
  "sheet": "Sheet1",
  "start_cell": "A1",
  "end_cell": "D20"
}
```

## Windows, WSL, and VS Code

If your `mcp.json` lives under a Windows path such as:

```text
C:\Users\<you>\AppData\Roaming\Code\User\mcp.json
```

VS Code usually starts the MCP server from Windows, not from WSL. In that case Windows must have
Node.js and Python installed, and Excel file paths should be Windows paths.

If you want VS Code on Windows to run the server inside WSL, call `wsl` explicitly:

```json
{
  "mcpServers": {
    "excel-tools-mcp": {
      "type": "stdio",
      "command": "wsl",
      "args": [
        "bash",
        "-lc",
        "npx --yes @wasziyang/excel-tools-mcp"
      ]
    }
  }
}
```

When the server runs in WSL, use Linux/WSL paths:

```json
{
  "file_path": "/mnt/c/Users/Alice/Documents/report.xlsx",
  "sheet": "Sheet1",
  "start_cell": "A1",
  "end_cell": "D20"
}
```

## Docker

Build locally:

```bash
docker build -t excel-tools-mcp .
```

MCP client config example:

```json
{
  "command": "docker",
  "args": [
    "run",
    "-i",
    "--rm",
    "-v",
    "/absolute/path/to/excel/files:/workspace",
    "excel-tools-mcp"
  ]
}
```

Inside Docker, pass file paths under `/workspace`, for example `/workspace/report.xlsx`.

Docker images are not published yet.

## Publishing Notes

Release checklist:

- Update versions in `pyproject.toml`, `package.json`, and the npm launcher cache directory.
- Build and publish the Python package to PyPI.
- Publish the npm launcher if the npx route should install the new version.

Basic npm publish flow:

```bash
npm login
npm publish --access public
```

Basic PyPI publish flow:

```bash
uv build
python3 -m twine upload dist/*
```
