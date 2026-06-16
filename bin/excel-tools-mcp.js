#!/usr/bin/env node

import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { homedir, platform } from "node:os";
import { spawn, spawnSync } from "node:child_process";

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const cacheRoot =
  process.env.EXCEL_TOOLS_MCP_CACHE_DIR ||
  join(homedir(), ".cache", "excel-tools-mcp");
const venvDir = join(cacheRoot, "venv-0.1.0");
const installMarker = join(venvDir, ".excel-tools-mcp-installed");
const isWindows = platform() === "win32";
const venvPython = join(venvDir, isWindows ? "Scripts/python.exe" : "bin/python");

function findPython() {
  const candidates = [
    process.env.EXCEL_TOOLS_MCP_PYTHON,
    "python3",
    "python",
  ].filter(Boolean);

  for (const candidate of candidates) {
    const result = spawnSync(candidate, ["--version"], { stdio: "ignore" });
    if (result.status === 0) {
      return candidate;
    }
  }

  console.error(
    "excel-tools-mcp requires Python 3.10+. Install Python or use the Docker image."
  );
  process.exit(1);
}

function runChecked(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    stdio: ["inherit", "inherit", "pipe"],
    ...options,
  });

  if (result.status !== 0) {
    if (result.stderr) {
      process.stderr.write(result.stderr);
    }
    process.exit(result.status ?? 1);
  }
}

function hasInstalledDependencies() {
  if (!existsSync(venvPython) || !existsSync(installMarker)) {
    return false;
  }

  const result = spawnSync(
    venvPython,
    ["-c", "import mcp, openpyxl, pydantic, dotenv"],
    { stdio: "ignore" }
  );
  return result.status === 0;
}

function ensurePip() {
  const pipCheck = spawnSync(venvPython, ["-m", "pip", "--version"], {
    stdio: "ignore",
  });
  if (pipCheck.status === 0) {
    return;
  }

  const ensurePip = spawnSync(venvPython, ["-m", "ensurepip", "--upgrade"], {
    encoding: "utf8",
    stdio: ["inherit", "inherit", "pipe"],
  });
  if (ensurePip.status === 0) {
    return;
  }

  if (ensurePip.stderr) {
    process.stderr.write(ensurePip.stderr);
  }
  console.error(
    "Python pip is not available in the virtual environment. On Debian/Ubuntu, install python3-pip and python3-venv, then remove ~/.cache/excel-tools-mcp and retry."
  );
  process.exit(ensurePip.status ?? 1);
}

function ensureVenv() {
  if (process.env.EXCEL_TOOLS_MCP_SKIP_INSTALL === "1" && existsSync(venvPython)) {
    return;
  }

  if (hasInstalledDependencies()) {
    return;
  }

  if (!existsSync(venvPython)) {
    const python = findPython();
    mkdirSync(cacheRoot, { recursive: true });
    const venvResult = spawnSync(python, ["-m", "venv", venvDir], {
      encoding: "utf8",
      stdio: ["inherit", "inherit", "pipe"],
    });

    if (venvResult.status !== 0) {
      if (venvResult.stderr) {
        process.stderr.write(venvResult.stderr);
      }
      console.error(
        "Failed to create a Python virtual environment. On Debian/Ubuntu, install python3-venv or use the Docker image."
      );
      process.exit(venvResult.status ?? 1);
    }
  }

  ensurePip();
  runChecked(venvPython, ["-m", "pip", "install", "--upgrade", "pip"]);
  runChecked(venvPython, ["-m", "pip", "install", packageRoot]);
  writeFileSync(installMarker, new Date().toISOString());
}

ensureVenv();

const child = spawn(venvPython, ["-m", "excel_tools.server"], {
  stdio: "inherit",
  env: process.env,
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 0);
});
