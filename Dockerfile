FROM python:3.12-slim

WORKDIR /app

COPY pyproject.toml README.md ./
COPY excel_tools ./excel_tools

RUN pip install --no-cache-dir .

WORKDIR /workspace

ENTRYPOINT ["excel-tools-mcp"]
