# WMI AI Wrapper - Project Overview

## Project Description

WMI AI Wrapper is a Python-based command-line interface (CLI) and AI-powered agent for Windows Management Instrumentation (WMI). It provides both traditional CLI commands and a natural language interface for Windows system administration tasks, making WMI accessible to both technical and non-technical users.

The project combines modern Python tooling (UV, Typer, Rich) with Microsoft Agent Framework to deliver a powerful yet user-friendly interface for querying and managing Windows systems.

## Key Features

### Core Capabilities
- **Complete WMI Access**: Execute any WQL query or WMI class operation
- **Pre-built Commands**: Ready-to-use commands for common tasks (services, processes, disks, network)
- **AI-Powered Agent**: Natural language interface using Microsoft Agent Framework
- **Dual AI Providers**: Support for both Ollama (local) and Azure OpenAI (cloud)
- **Administrator Detection**: Automatic privilege detection and handling
- **Multiple Output Formats**: Rich tables or JSON for scripting
- **Python API**: Programmatic access for custom integrations

### Technical Features
- Built with modern Python 3.8+ and UV package manager
- Rich terminal UI with formatted tables and colors
- Type-safe CLI with Typer framework
- Streaming responses for AI agent interactions
- Environment-based configuration

## Vision and Scope

### Vision
Make Windows system administration accessible through both traditional CLI commands and conversational AI, bridging the gap between expert administrators and those learning Windows management.

### Scope

**In Scope:**
- Windows system information queries (OS, hardware, BIOS)
- Resource monitoring (CPU, memory, disk usage)
- Service and process management
- Network configuration inspection
- Custom WQL query execution
- Natural language system queries via AI agent
- Both local (Ollama) and cloud (Azure OpenAI) AI models

**Out of Scope:**
- Non-Windows operating systems
- Real-time system modification (focused on querying and monitoring)
- Web UI or GUI interfaces
- Vendor-specific hardware management (Dell, HP, etc.)

## Architecture

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                         User Interface                       │
├──────────────────────┬──────────────────────────────────────┤
│   CLI Commands       │        AI Agent (Natural Language)   │
│   (wmi-cli)          │        (wmi-agent)                   │
└──────────┬───────────┴──────────────────┬───────────────────┘
           │                               │
           │                               │
           ▼                               ▼
┌──────────────────────┐      ┌───────────────────────────────┐
│   WMI Wrapper        │      │   Microsoft Agent Framework   │
│   (wmi_wrapper.py)   │      │   (agent.py)                  │
└──────────┬───────────┘      └───────────┬───────────────────┘
           │                               │
           │                   ┌───────────┴────────────┐
           │                   │                        │
           │                   ▼                        ▼
           │          ┌─────────────────┐    ┌──────────────────┐
           │          │  Ollama Server  │    │  Azure OpenAI    │
           │          │  (Local LLM)    │    │  (Cloud)         │
           │          └─────────────────┘    └──────────────────┘
           │                   │                        │
           │                   └───────────┬────────────┘
           │                               │
           │                               ▼
           │                   ┌───────────────────────────────┐
           │                   │   WMI Tools (wmi_tools.py)    │
           │                   │   - @ai_function decorated    │
           │                   │   - 12 WMI operations         │
           │                   └───────────┬───────────────────┘
           │                               │
           └───────────────────────────────┘
                               │
                               ▼
                   ┌───────────────────────┐
                   │   Windows WMI Layer   │
                   │   (System APIs)       │
                   └───────────────────────┘
```

### Component Interaction Flow

**CLI Command Flow:**
```
User → CLI Command → WMI Wrapper → Windows WMI → Response → Formatted Output
```

**AI Agent Flow:**
```
User Query (Natural Language)
    ↓
Agent Framework (Ollama/Azure OpenAI)
    ↓
LLM analyzes query & selects appropriate WMI tool
    ↓
WMI Tools execute via WMI Wrapper
    ↓
Windows WMI returns data
    ↓
Agent formats & streams response to user
```

## Example Workflows and Use Cases

### Use Case 1: System Information Query (CLI)

**Scenario:** Administrator needs quick system specs

**Input:**
```powershell
wmi-cli system-info
```

**Output:**
```
                 System Information                 
┌───────────────────────┬──────────────────────────┐
│ Computer Name         │ DESKTOP-F2F6MAB          │
│ Manufacturer          │ Dell Inc.                │
│ Model                 │ Latitude 9520            │
│ OS Name               │ Microsoft Windows 11 Pro │
│ OS Version            │ 10.0.26100               │
│ OS Architecture       │ 64-bit                   │
│ System Type           │ x64-based PC             │
│ Total Physical Memory │ 15.73 GB                 │
│ BIOS Version          │ DELL   - 2               │
│ Serial Number         │ 8G921J3                  │
└───────────────────────┴──────────────────────────┘
```

### Use Case 2: Service Management (CLI)

**Scenario:** Check running services

**Input:**
```powershell
wmi-cli services --state Running --start-mode Auto
```

**Output:**
```
Services (Filtered: state=Running, start_mode=Auto)
┌──────────────────────┬─────────────┬──────────┬────────────┐
│ Name                 │ Display Name│ State    │ Start Mode │
├──────────────────────┼─────────────┼──────────┼────────────┤
│ wuauserv             │ Windows...  │ Running  │ Auto       │
│ Spooler              │ Print...    │ Running  │ Auto       │
└──────────────────────┴─────────────┴──────────┴────────────┘
```

### Use Case 3: Natural Language Query (AI Agent)

**Scenario:** Non-technical user wants memory information

**Input:**
```
You: How much memory does my computer have and how much is being used?
```

**Agent Processing:**
1. Understands user intent (memory information)
2. Calls `get_memory_info()` WMI tool
3. Formats response conversationally

**Output:**
```
Your computer has 15.73 GB of total physical memory.

Current memory usage:
• Used: 8.42 GB (53.5%)
• Available: 7.31 GB (46.5%)

Your system has a moderate amount of memory available. 
If you're experiencing slowness, you might want to close 
some applications to free up more memory.
```

### Use Case 4: Process Monitoring (AI Agent)

**Scenario:** Identify memory-intensive processes

**Input:**
```
You: What processes are using the most memory?
```

**Agent Processing:**
1. Calls `list_processes()` tool
2. Sorts by memory usage
3. Presents top consumers

**Output:**
```
Here are the top memory-consuming processes on your system:

1. chrome.exe - 1,234 MB (12 instances)
2. code.exe - 856 MB (Visual Studio Code)
3. Teams.exe - 643 MB (Microsoft Teams)
4. firefox.exe - 521 MB
5. explorer.exe - 287 MB (Windows Explorer)

Chrome is using the most memory with multiple browser tabs open. 
Consider closing unused tabs to free up memory.
```

### Use Case 5: Custom WQL Query (CLI)

**Scenario:** Advanced user needs specific data

**Input:**
```powershell
wmi-cli query "SELECT Caption, FreeSpace, Size FROM Win32_LogicalDisk WHERE DriveType=3"
```

**Output:**
```
Query Results (3 objects returned)
┌─────────┬────────────┬─────────────┐
│ Caption │ FreeSpace  │ Size        │
├─────────┼────────────┼─────────────┤
│ C:      │ 125.4 GB   │ 475.9 GB    │
│ D:      │ 850.2 GB   │ 1863.0 GB   │
└─────────┴────────────┴─────────────┘
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
└── README.md             # User documentation
```

### Key Components

- **`wmi_cli/cli.py`**: Typer-based CLI with 10 commands for WMI operations
- **`wmi_cli/wmi_wrapper.py`**: Core WMI abstraction layer using Python's `wmi` library
- **`wmi_cli/modules.py`**: Specialized managers (ServiceManager, ProcessManager, etc.)
- **`agent.py`**: WMIAgent class using Microsoft Agent Framework
- **`cli_agent.py`**: Interactive terminal interface for the AI agent
- **`wmi_tools.py`**: 12 @ai_function decorated tools for agent use

## Technology Stack

### Core Technologies
- **Python 3.8+**: Primary programming language
- **UV**: Modern Python package manager
- **WMI Library**: Python wrapper for Windows WMI
- **Typer**: CLI framework
- **Rich**: Terminal formatting and output

### AI/Agent Technologies
- **Microsoft Agent Framework**: Core agent orchestration
- **Ollama**: Local LLM inference server
- **Azure OpenAI**: Cloud-based AI models
- **Python-dotenv**: Environment configuration

### Development Tools
- **Black**: Code formatting
- **Ruff**: Fast Python linter
- **MyPy**: Static type checking

## Configuration

The project supports flexible configuration through environment variables:

### CLI Configuration
- No configuration required for basic WMI operations
- Runs with default Windows credentials

### Agent Configuration
```bash
# Choose provider
AGENT_PROVIDER=ollama  # or "azure"

# Ollama (Local)
OLLAMA_MODEL=gpt-oss:120b
OLLAMA_ENDPOINT=http://localhost:11434/v1

# Azure OpenAI (Cloud)
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_API_KEY=your-key
AZURE_OPENAI_DEPLOYMENT=gpt-4o
AZURE_OPENAI_API_VERSION=2024-08-01-preview
```

## Target Users

1. **Windows System Administrators**: Traditional CLI users who need quick, scriptable access to WMI
2. **IT Support Teams**: Technical staff who prefer natural language queries
3. **DevOps Engineers**: Automation specialists needing programmatic WMI access
4. **Students/Learners**: Those learning Windows administration through conversational interface
5. **Python Developers**: Engineers building Windows management tools

## Future Enhancements

- Additional WMI namespaces support (security, performance counters)
- Scheduled monitoring and alerting capabilities
- Export/import configuration profiles
- Web dashboard for remote management
- Plugin system for custom WMI tools
- Multi-system management (query multiple computers)
- Enhanced streaming for long-running operations
