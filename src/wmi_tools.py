"""
WMI Tools for Agent Framework

This module provides WMI functionality as agent tools using the @ai_function decorator.
These tools allow the agent to interact with Windows Management Instrumentation.
"""

from typing import Annotated, Optional, List
from pydantic import Field
from agent_framework import ai_function
from .wmi_cli.wmi_wrapper import WMIWrapper, is_admin, format_bytes
from .wmi_cli.modules import SystemMonitor, ProcessManager


# Global WMI instances (initialized once)
_wrapper = None
_system_mon = None
_process_mgr = None


def _init_wmi():
    """Initialize WMI instances (called lazily on first use)"""
    global _wrapper, _system_mon, _process_mgr
    if _wrapper is None:
        _wrapper = WMIWrapper()
        _system_mon = SystemMonitor()
        _process_mgr = ProcessManager()


# WMI Tool Functions decorated with @ai_function

@ai_function(description="Get detailed system information including OS, hardware, and BIOS details")
def get_system_info() -> str:
    """Retrieves comprehensive system information."""
    try:
        _init_wmi()
        os_info = _wrapper.get_operating_system()
        cs_info = _wrapper.get_computer_system()
        bios_info = _wrapper.get_bios()
        
        result = "System Information:\n"
        result += f"  OS: {os_info.Caption} {os_info.Version}\n"
        result += f"  Computer: {cs_info.Name}\n"
        result += f"  Manufacturer: {cs_info.Manufacturer}\n"
        result += f"  Model: {cs_info.Model}\n"
        result += f"  Architecture: {os_info.OSArchitecture}\n"
        result += f"  Memory: {format_bytes(int(cs_info.TotalPhysicalMemory))}\n"
        result += f"  BIOS: {bios_info.Version if bios_info else 'N/A'}\n"
        result += f"  Serial Number: {bios_info.SerialNumber if bios_info else 'N/A'}\n"
        
        return result
    except Exception as e:
        return f"Error getting system info: {str(e)}"


@ai_function(description="Get memory usage information including total, used, and available memory")
def get_memory_info() -> str:
    """Retrieves current memory usage statistics."""
    try:
        _init_wmi()
        info = _system_mon.get_memory_info()
        
        result = "Memory Information:\n"
        result += f"  Total: {format_bytes(info['total_bytes'])}\n"
        result += f"  Used: {format_bytes(info['used_bytes'])}\n"
        result += f"  Free: {format_bytes(info['free_bytes'])}\n"
        result += f"  Usage: {info['used_percentage']:.1f}%\n"
        
        return result
    except Exception as e:
        return f"Error getting memory info: {str(e)}"


@ai_function(description="Get CPU usage and processor information")
def get_cpu_info() -> str:
    """Retrieves CPU usage and processor details."""
    try:
        _init_wmi()
        # Use direct WMI query since get_cpu_info() returns schema, not data
        cpu_list = _wrapper.query("SELECT * FROM Win32_Processor")
        
        if not cpu_list:
            return "No CPU information available"
        
        cpu = cpu_list[0]
        
        result = "CPU Information:\n"
        result += f"  Processor: {cpu.Name}\n"
        result += f"  Manufacturer: {cpu.Manufacturer}\n"
        result += f"  Cores: {cpu.NumberOfCores}\n"
        result += f"  Logical Processors: {cpu.NumberOfLogicalProcessors}\n"
        result += f"  Max Speed: {cpu.MaxClockSpeed} MHz\n"
        result += f"  Current Speed: {cpu.CurrentClockSpeed} MHz\n"
        if cpu.LoadPercentage is not None:
            result += f"  Load: {cpu.LoadPercentage}%\n"
        
        return result
    except Exception as e:
        return f"Error getting CPU info: {str(e)}"


@ai_function(description="Get disk drive information including size, free space, and usage")
def get_disk_info() -> str:
    """Retrieves information about all disk drives."""
    try:
        _init_wmi()
        # Query logical disks directly (DriveType=3 means local disk)
        disks = _wrapper.query("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3")
        
        result = "Disk Drives:\n"
        for disk in disks:
            if disk.Size:
                free_space = int(disk.FreeSpace) if disk.FreeSpace else 0
                total_space = int(disk.Size)
                used_pct = ((total_space - free_space) / total_space * 100) if total_space > 0 else 0
                
                result += f"  {disk.Name}:\n"
                result += f"    Size: {format_bytes(total_space)}\n"
                result += f"    Free: {format_bytes(free_space)}\n"
                result += f"    Used: {used_pct:.1f}%\n"
                result += f"    File System: {disk.FileSystem if disk.FileSystem else 'N/A'}\n"
        
        return result if result != "Disk Drives:\n" else "No disk information available"
    except Exception as e:
        return f"Error getting disk info: {str(e)}"


@ai_function(description="Get network adapter configuration and IP addresses")
def get_network_info() -> str:
    """Retrieves network adapter configuration."""
    try:
        _init_wmi()
        # Query network adapters directly (IPEnabled=True means active adapters)
        adapters = _wrapper.query("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
        
        if not adapters:
            return "No active network adapters found"
        
        result = "Network Adapters:\n"
        for adapter in adapters:
            result += f"  {adapter.Description}:\n"
            result += f"    MAC: {adapter.MACAddress if adapter.MACAddress else 'N/A'}\n"
            if adapter.IPAddress:
                result += f"    IP: {', '.join(adapter.IPAddress)}\n"
            if adapter.DefaultIPGateway:
                result += f"    Gateway: {', '.join(adapter.DefaultIPGateway)}\n"
            result += f"    DHCP Enabled: {'Yes' if adapter.DHCPEnabled else 'No'}\n"
        
        return result
    except Exception as e:
        return f"Error getting network info: {str(e)}"


@ai_function(description="Get system uptime since last boot")
def get_uptime() -> str:
    """Retrieves system uptime information."""
    try:
        _init_wmi()
        uptime_info = _system_mon.get_uptime()
        
        result = "System Uptime:\n"
        result += f"  Last Boot: {uptime_info.get('last_boot_time', 'N/A')}\n"
        result += f"  Uptime: {uptime_info.get('uptime_days', 0)} days, "
        result += f"{uptime_info.get('uptime_hours', 0)} hours, "
        result += f"{uptime_info.get('uptime_minutes', 0)} minutes\n"
        
        return result
    except Exception as e:
        return f"Error getting uptime: {str(e)}"


@ai_function(description="Check if running with administrator privileges")
def check_admin_privileges() -> str:
    """Checks if the current process has administrator privileges."""
    try:
        if is_admin():
            return "Running with Administrator privileges"
        else:
            return "NOT running with Administrator privileges. Some operations may be restricted."
    except Exception as e:
        return f"Error checking admin privileges: {str(e)}"


@ai_function(description="List Windows services with optional filtering")
def list_services(
    state: Annotated[Optional[str], Field(description="Filter by state: 'Running' or 'Stopped'")] = None
) -> str:
    """Lists Windows services, optionally filtered by state."""
    try:
        _init_wmi()
        # ServiceManager doesn't have get_services, use wrapper directly
        if state:
            query = f"SELECT Name, DisplayName, State, StartMode, Status FROM Win32_Service WHERE State = '{state}'"
        else:
            query = "SELECT Name, DisplayName, State, StartMode, Status FROM Win32_Service"
        
        services = _wrapper.query(query)
        
        result = f"Windows Services ({len(services)}):\n"
        for i, svc in enumerate(services[:20], 1):  # Limit to first 20
            result += f"  {i}. {svc.Name} - {svc.State}\n"
            result += f"     Display: {svc.DisplayName}\n"
        
        if len(services) > 20:
            result += f"\n  ... and {len(services) - 20} more services\n"
        
        return result
    except Exception as e:
        return f"Error listing services: {str(e)}"


@ai_function(description="Get status of a specific Windows service")
def get_service_status(
    service_name: Annotated[str, Field(description="Name of the service to query")]
) -> str:
    """Gets the status of a specific Windows service."""
    try:
        _init_wmi()
        # Query service directly using WMI
        query = f"SELECT Name, DisplayName, State, StartMode, Status FROM Win32_Service WHERE Name = '{service_name}'"
        services = _wrapper.query(query)
        
        if services:
            svc = services[0]
            result = f"Service: {svc.Name}\n"
            result += f"  Display Name: {svc.DisplayName}\n"
            result += f"  State: {svc.State}\n"
            result += f"  Start Mode: {svc.StartMode}\n"
            result += f"  Status: {svc.Status}\n"
            return result
        else:
            return f"Service '{service_name}' not found"
    except Exception as e:
        return f"Error getting service status: {str(e)}"


@ai_function(description="List running processes")
def list_processes() -> str:
    """Lists currently running processes."""
    try:
        _init_wmi()
        # Use get_high_memory_processes with min_memory_mb parameter (default 100MB)
        # Lower the threshold to 0 to get all processes, then take top 15
        processes = _process_mgr.get_high_memory_processes(min_memory_mb=0)
        
        if not processes:
            return "No processes found"
        
        # Sort by memory and take top 15
        processes_sorted = sorted(processes, key=lambda p: int(p.WorkingSetSize) if p.WorkingSetSize else 0, reverse=True)[:15]
        
        result = f"Running Processes (Top 15 by Memory):\n"
        for i, proc in enumerate(processes_sorted, 1):
            mem_bytes = int(proc.WorkingSetSize) if proc.WorkingSetSize else 0
            mem_mb = mem_bytes / (1024 * 1024)
            result += f"  {i}. {proc.Name} (PID: {proc.ProcessId})\n"
            result += f"     Memory: {mem_mb:.1f} MB\n"
        
        return result
    except Exception as e:
        return f"Error listing processes: {str(e)}"


@ai_function(description="Get process CPU and memory usage with performance metrics")
def get_process_performance() -> str:
    """Gets CPU and memory usage for top processes using performance counters."""
    try:
        _init_wmi()
        
        # Query Win32_PerfFormattedData_PerfProc_Process for CPU percentages
        perf_query = "SELECT Name, IDProcess, PercentProcessorTime, WorkingSet FROM Win32_PerfFormattedData_PerfProc_Process WHERE Name != '_Total' AND Name != 'Idle'"
        perf_processes = _wrapper.query(perf_query)
        
        if not perf_processes:
            return "No performance data available. This can happen if performance counters are disabled or need to be rebuilt."
        
        # Sort by CPU usage and take top 15
        valid_procs = [p for p in perf_processes if p.PercentProcessorTime is not None and p.IDProcess != 0]
        procs_sorted = sorted(valid_procs, key=lambda p: int(p.PercentProcessorTime) if p.PercentProcessorTime else 0, reverse=True)[:15]
        
        result = "Process Performance (Top 15 by CPU Usage):\n\n"
        for i, proc in enumerate(procs_sorted, 1):
            cpu_percent = int(proc.PercentProcessorTime) if proc.PercentProcessorTime else 0
            mem_bytes = int(proc.WorkingSet) if proc.WorkingSet else 0
            mem_mb = mem_bytes / (1024 * 1024)
            pid = proc.IDProcess
            
            result += f"  {i}. {proc.Name} (PID: {pid})\n"
            result += f"     CPU: {cpu_percent}%\n"
            result += f"     Memory: {mem_mb:.1f} MB\n\n"
        
        return result
    except Exception as e:
        return f"Error getting process performance: {str(e)}\n\nNote: Performance counters may not be available. Try using 'list_processes' for memory-based process listing instead."


@ai_function(description="Execute a custom WQL query")
def execute_wql_query(
    query: Annotated[str, Field(description="WQL query to execute (e.g., 'SELECT * FROM Win32_Service')")]
) -> str:
    """Executes a custom WQL query and returns results."""
    try:
        _init_wmi()
        results = _wrapper.query(query)
        
        if not results:
            return "Query returned no results"
        
        # Return first 5 results
        result = f"Query Results ({len(results)} total, showing first 5):\n"
        for i, item in enumerate(results[:5], 1):
            result += f"\n  Result {i}:\n"
            # Show first few properties
            props = list(item.properties.keys())[:10]
            for prop in props:
                try:
                    value = getattr(item, prop)
                    if value is not None:
                        result += f"    {prop}: {value}\n"
                except:
                    pass
        
        return result
    except Exception as e:
        return f"Error executing query: {str(e)}"


def get_wmi_tools() -> list:
    """
    Returns a list of all WMI tool functions for the agent.
    
    Returns:
        List of @ai_function decorated functions
    """
    return [
        get_system_info,
        get_memory_info,
        get_cpu_info,
        get_disk_info,
        get_network_info,
        get_uptime,
        check_admin_privileges,
        list_services,
        get_service_status,
        list_processes,
        get_process_performance,
        execute_wql_query,
    ]
