# WMI CLI - Windows Management Instrumentation Wrapper

A Python CLI wrapper for Windows Management Instrumentation (WMI) built with UV. Execute any WMI query or operation from the command line with a simple, modern interface.

## Features

- 🔧 Complete WMI access - execute any query or class operation
- 🎯 Pre-built commands for common tasks (services, processes, disks, network)
- 🤖 **AI Agent** - Natural language interface for WMI using Microsoft Agent Framework
- 👑🔐 Administrator privilege detection and handling
- 📦 Modern tooling with UV, Typer, and Rich
- 🐍 Python API for programmatic use

## Prerequisites

- Windows OS (WMI is Windows-specific)
- Python 3.10+
- [UV Package Manager](https://github.com/astral-sh/uv)

> **Note:** This project uses UV for package management. All commands must be prefixed with `uv run` (e.g., `uv run wmi-cli`).

## Quick Start

```powershell
# 1. Install UV
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# 2. Clone and navigate to project
cd wmi-ai-wrapper

# 3. Sync dependencies
uv sync

# 4. Try it out
uv run wmi-cli --help
uv run wmi-cli system-info
uv run wmi-cli services --state Running
```

## WMI Agent (AI Interface)

The WMI Agent provides a natural language interface for Windows system management using AI.

### Quick Start with Agent

```powershell
# Sync dependencies (includes agent extras)
uv sync

# Run with Ollama (local)
ollama pull gpt-oss:120b
ollama serve
uv run wmi-agent

# Run with Azure OpenAI
$env:AGENT_PROVIDER="azure"
$env:AZURE_OPENAI_ENDPOINT="https://your-resource.openai.azure.com"
$env:AZURE_OPENAI_API_KEY="your-key"
$env:AZURE_OPENAI_DEPLOYMENT="gpt-4o"
uv run wmi-agent
```

### Example Agent Queries

```
You: What's my current memory usage?
You: Show me running services
You: List disk drives and their space
You: What processes are using the most memory?
You: Get system uptime
```

## Available Commands

### General WMI Commands

| Command | Description |
|---------|-------------|
| `wmi-cli admin-check` | Check if running with admin privileges |
| `wmi-cli system-info` | Display system information |
| `wmi-cli services` | List and filter Windows services |
| `wmi-cli processes` | List running processes |
| `wmi-cli disks` | Show disk drives and usage |
| `wmi-cli network` | Display network adapter configuration |
| `wmi-cli list-classes` | List all available WMI classes |
| `wmi-cli class-info <name>` | Get properties of a WMI class |
| `wmi-cli query "<WQL>"` | Execute raw WQL queries |

### Command Examples

```powershell
# System info
uv run wmi-cli system-info
uv run wmi-cli system-info --output-format json

# Services
uv run wmi-cli services --state Running
uv run wmi-cli services --start-mode Auto --state Stopped

# Processes
uv run wmi-cli processes
uv run wmi-cli processes --name "chrome.exe"

# Disks
uv run wmi-cli disks
uv run wmi-cli disks --drive-type 3  # Local disks only

# Network
uv run wmi-cli network --output-format json

# WMI Classes
uv run wmi-cli list-classes --filter-text "Win32"
uv run wmi-cli class-info Win32_Service

# Raw queries
uv run wmi-cli query "SELECT * FROM Win32_OperatingSystem"
```

## Python API

Use the library programmatically in your scripts:

```python
from src.wmi_cli.wmi_wrapper import WMIWrapper
from src.wmi_cli.modules import SystemMonitor, ServiceManager

# Basic operations
wrapper = WMIWrapper()
services = wrapper.get_services(State="Running")

# System monitoring
monitor = SystemMonitor()
memory = monitor.get_memory_info()
uptime = monitor.get_uptime()
print(f"Memory: {memory['used_percentage']:.1f}% used")
print(f"Uptime: {uptime['uptime_days']} days")

# Service management
mgr = ServiceManager()
status = mgr.get_service_status("wuauserv")
```

**Available modules:** ServiceManager, ProcessManager, SystemMonitor, NetworkManager, EventLogReader, HardwareInfo, SecurityManager

## Administrator Privileges

Some operations require admin rights (start/stop services, terminate processes, read security logs). Run PowerShell as Administrator for these operations.

Check your current privileges:
```powershell
wmi-cli admin-check
```

## Output Formats

All commands support `--output-format`:
- `table` (default): Rich formatted tables
- `json`: JSON for scripting/parsing

```powershell
wmi-cli services --output-format json > services.json
```

## Project Structure

```
wmi-ai-wrapper/
├── src/                  # All source code
│   ├── wmi_cli/          # Main CLI package
│   │   ├── cli.py        # CLI commands
│   │   ├── wmi_wrapper.py    # Core WMI wrapper
│   │   └── modules.py    # Specialized modules
│   ├── agent.py          # AI Agent implementation
│   ├── cli_agent.py      # Agent CLI interface
│   └── wmi_tools.py      # WMI tools for agent
├── .env.example          # Example environment variables
├── pyproject.toml        # Project configuration
└── README.md             # This file
```

## Development

```powershell
# Install dev dependencies
uv pip install -e ".[dev]"

# Format code
black src/
ruff check src/
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Module not found" | Activate venv: `.venv\Scripts\activate` |
| "Permission denied" | Run PowerShell as Administrator |
| pywin32 issues | `uv pip install --force-reinstall pywin32` |
| WMI timeout | Check network, firewall, and WMI service status |

## Resources

- [WMI Reference](https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-reference) - Microsoft WMI documentation
- [tjguk/wmi](https://github.com/tjguk/wmi) - Python WMI library
- [WQL Reference](https://learn.microsoft.com/en-us/windows/win32/wmisdk/wql-sql-for-wmi) - Query language docs

## License

MIT License

## Acknowledgments

Built with [tjguk/wmi](https://github.com/tjguk/wmi), [UV](https://github.com/astral-sh/uv), [Typer](https://typer.tiangolo.com/), and [Rich](https://github.com/Textualize/rich)
