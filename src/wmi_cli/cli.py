"""
Command-line interface for WMI operations.
"""
import json
import sys
from typing import Optional, List

import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.syntax import Syntax

from .wmi_wrapper import WMIWrapper, is_admin, format_bytes, wmi_object_to_dict

app = typer.Typer(
    name="wmi-cli",
    help="Windows Management Instrumentation (WMI) CLI wrapper",
    add_completion=True,
)
console = Console()


@app.command()
def query(
    wql: str = typer.Argument(..., help="WQL query to execute"),
    namespace: str = typer.Option("root\\cimv2", help="WMI namespace"),
    output_format: str = typer.Option("table", help="Output format: table, json, raw"),
    computer: str = typer.Option(".", help="Computer name (. for local)"),
):
    """
    Execute a raw WQL query.
    
    Note: Due to WMI library limitations, raw queries work best with SELECT * queries.
    For specific property queries, consider using the built-in commands
    (services, processes, system-info, etc.) instead.
    """
    try:
        wrapper = WMIWrapper(computer=computer, namespace=namespace)
        results = wrapper.query(wql)
        
        if not results:
            console.print("[yellow]No results found[/yellow]")
            return
        
        if output_format == "json":
            # Try to extract real property values
            output = []
            for obj in results:
                try:
                    # Try to get actual WMI properties
                    item_dict = {}
                    for prop_name in obj.properties:
                        try:
                            item_dict[prop_name] = getattr(obj, prop_name)
                        except:
                            item_dict[prop_name] = None
                    output.append(item_dict)
                except:
                    # Fallback to standard conversion
                    output.append(wmi_object_to_dict(obj))
            console.print(json.dumps(output, indent=2, default=str))
        elif output_format == "table":
            # For table output, show actual properties
            if results:
                try:
                    # Get actual properties from first object
                    first_obj = results[0]
                    properties = [p for p in first_obj.properties][:10]  # Limit to 10 columns
                    
                    table = Table(title=f"Query Results")
                    for prop in properties:
                        table.add_column(prop, style="cyan")
                    
                    for obj in results:
                        row = []
                        for prop in properties:
                            try:
                                value = getattr(obj, prop)
                                row.append(str(value) if value is not None else "N/A")
                            except:
                                row.append("N/A")
                        table.add_row(*row)
                    
                    console.print(table)
                except:
                    # Fallback to old method
                    _display_table(results, f"Query Results: {wql}")
        else:
            for obj in results:
                console.print(obj)
                
    except Exception as e:
        console.print(f"[red]Error executing query: {e}[/red]")
        import traceback
        traceback.print_exc()
        raise typer.Exit(1)


@app.command()
def list_classes(
    namespace: str = typer.Option("root\\cimv2", help="WMI namespace"),
    filter_text: Optional[str] = typer.Option(None, help="Filter class names"),
):
    """List all available WMI classes."""
    try:
        wrapper = WMIWrapper(namespace=namespace)
        classes = wrapper.list_classes()
        
        if filter_text:
            classes = [c for c in classes if filter_text.lower() in c.lower()]
        
        classes.sort()
        
        table = Table(title=f"WMI Classes in {namespace}")
        table.add_column("Class Name", style="cyan")
        
        for cls in classes:
            table.add_row(cls)
        
        console.print(table)
        console.print(f"\n[green]Total: {len(classes)} classes[/green]")
        
    except Exception as e:
        console.print(f"[red]Error listing classes: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def class_info(
    class_name: str = typer.Argument(..., help="WMI class name"),
    namespace: str = typer.Option("root\\cimv2", help="WMI namespace"),
):
    """Get information about a WMI class (properties and sample instance)."""
    try:
        wrapper = WMIWrapper(namespace=namespace)
        properties = wrapper.get_class_properties(class_name)
        
        if not properties:
            console.print(f"[yellow]No instances found for class {class_name}[/yellow]")
            return
        
        table = Table(title=f"Properties of {class_name}")
        table.add_column("Property", style="cyan")
        
        for prop in sorted(properties):
            table.add_row(prop)
        
        console.print(table)
        console.print(f"\n[green]Total: {len(properties)} properties[/green]")
        
        # Show a sample instance
        instances = wrapper.get_class(class_name)
        if instances:
            console.print("\n[bold]Sample Instance:[/bold]")
            sample_data = wmi_object_to_dict(instances[0])
            console.print(json.dumps(sample_data, indent=2, default=str))
        
    except Exception as e:
        console.print(f"[red]Error getting class info: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def services(
    name: Optional[str] = typer.Option(None, help="Filter by service name"),
    state: Optional[str] = typer.Option(None, help="Filter by state (Running, Stopped, etc.)"),
    start_mode: Optional[str] = typer.Option(None, help="Filter by start mode (Auto, Manual, etc.)"),
    output_format: str = typer.Option("table", help="Output format: table, json"),
):
    """List Windows services."""
    try:
        wrapper = WMIWrapper()
        filters = {}
        if name:
            filters["Name"] = name
        if state:
            filters["State"] = state
        if start_mode:
            filters["StartMode"] = start_mode
        
        results = wrapper.get_services(**filters)
        
        if not results:
            console.print("[yellow]No services found matching criteria[/yellow]")
            return
        
        if output_format == "json":
            output = [wmi_object_to_dict(svc, ["Name", "DisplayName", "State", "StartMode", "Status"]) 
                     for svc in results]
            console.print(json.dumps(output, indent=2, default=str))
        else:
            table = Table(title="Windows Services")
            table.add_column("Name", style="cyan")
            table.add_column("Display Name", style="white")
            table.add_column("State", style="green")
            table.add_column("Start Mode", style="yellow")
            table.add_column("Status", style="magenta")
            
            for svc in results:
                table.add_row(
                    svc.Name,
                    svc.DisplayName,
                    svc.State,
                    svc.StartMode,
                    svc.Status,
                )
            
            console.print(table)
            console.print(f"\n[green]Total: {len(results)} services[/green]")
            
    except Exception as e:
        console.print(f"[red]Error listing services: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def processes(
    name: Optional[str] = typer.Option(None, help="Filter by process name"),
    output_format: str = typer.Option("table", help="Output format: table, json"),
):
    """List running processes."""
    try:
        wrapper = WMIWrapper()
        filters = {}
        if name:
            filters["Name"] = name
        
        results = wrapper.get_processes(**filters)
        
        if not results:
            console.print("[yellow]No processes found[/yellow]")
            return
        
        if output_format == "json":
            output = [wmi_object_to_dict(proc, ["Name", "ProcessId", "ThreadCount", "WorkingSetSize", "CommandLine"]) 
                     for proc in results]
            console.print(json.dumps(output, indent=2, default=str))
        else:
            table = Table(title="Running Processes")
            table.add_column("PID", style="cyan")
            table.add_column("Name", style="white")
            table.add_column("Threads", style="yellow")
            table.add_column("Memory", style="green")
            
            for proc in results:
                memory = format_bytes(int(proc.WorkingSetSize)) if proc.WorkingSetSize else "N/A"
                table.add_row(
                    str(proc.ProcessId),
                    proc.Name,
                    str(proc.ThreadCount) if proc.ThreadCount else "N/A",
                    memory,
                )
            
            console.print(table)
            console.print(f"\n[green]Total: {len(results)} processes[/green]")
            
    except Exception as e:
        console.print(f"[red]Error listing processes: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def system_info(
    output_format: str = typer.Option("table", help="Output format: table, json"),
):
    """Display system information."""
    try:
        wrapper = WMIWrapper()
        os_info = wrapper.get_operating_system()
        cs_info = wrapper.get_computer_system()
        bios_info = wrapper.get_bios()
        
        if output_format == "json":
            output = {
                "operating_system": wmi_object_to_dict(os_info),
                "computer_system": wmi_object_to_dict(cs_info),
                "bios": wmi_object_to_dict(bios_info),
            }
            console.print(json.dumps(output, indent=2, default=str))
        else:
            table = Table(title="System Information", show_header=False)
            table.add_column("Property", style="cyan bold")
            table.add_column("Value", style="white")
            
            table.add_row("Computer Name", cs_info.Name)
            table.add_row("Manufacturer", cs_info.Manufacturer)
            table.add_row("Model", cs_info.Model)
            table.add_row("OS Name", os_info.Caption)
            table.add_row("OS Version", os_info.Version)
            table.add_row("OS Architecture", os_info.OSArchitecture)
            table.add_row("System Type", cs_info.SystemType)
            table.add_row("Total Physical Memory", format_bytes(int(cs_info.TotalPhysicalMemory)))
            table.add_row("BIOS Version", bios_info.Version if bios_info else "N/A")
            table.add_row("Serial Number", bios_info.SerialNumber if bios_info else "N/A")
            
            console.print(table)
            
    except Exception as e:
        console.print(f"[red]Error getting system info: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def disks(
    drive_type: Optional[int] = typer.Option(None, help="Filter by drive type (3=Local, 4=Network, 5=CD-ROM)"),
    output_format: str = typer.Option("table", help="Output format: table, json"),
):
    """List disk drives."""
    try:
        wrapper = WMIWrapper()
        filters = {}
        if drive_type is not None:
            filters["DriveType"] = drive_type
        
        results = wrapper.get_logical_disks(**filters)
        
        if not results:
            console.print("[yellow]No disks found[/yellow]")
            return
        
        if output_format == "json":
            output = [wmi_object_to_dict(disk, ["DeviceID", "VolumeName", "DriveType", "FileSystem", 
                                                "Size", "FreeSpace"]) for disk in results]
            console.print(json.dumps(output, indent=2, default=str))
        else:
            table = Table(title="Disk Drives")
            table.add_column("Drive", style="cyan")
            table.add_column("Label", style="white")
            table.add_column("Type", style="yellow")
            table.add_column("File System", style="green")
            table.add_column("Size", style="blue")
            table.add_column("Free Space", style="magenta")
            table.add_column("Used %", style="red")
            
            drive_types = {
                0: "Unknown",
                1: "No Root",
                2: "Removable",
                3: "Local",
                4: "Network",
                5: "CD-ROM",
                6: "RAM Disk"
            }
            
            for disk in results:
                size = int(disk.Size) if disk.Size else 0
                free = int(disk.FreeSpace) if disk.FreeSpace else 0
                used_pct = ((size - free) / size * 100) if size > 0 else 0
                
                table.add_row(
                    disk.DeviceID,
                    disk.VolumeName or "",
                    drive_types.get(disk.DriveType, "Unknown"),
                    disk.FileSystem or "N/A",
                    format_bytes(size) if size > 0 else "N/A",
                    format_bytes(free) if free > 0 else "N/A",
                    f"{used_pct:.1f}%" if size > 0 else "N/A",
                )
            
            console.print(table)
            console.print(f"\n[green]Total: {len(results)} disks[/green]")
            
    except Exception as e:
        console.print(f"[red]Error listing disks: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def network(
    output_format: str = typer.Option("table", help="Output format: table, json"),
):
    """Display network adapter configuration."""
    try:
        wrapper = WMIWrapper()
        results = wrapper.get_network_adapters(IPEnabled=True)
        
        if not results:
            console.print("[yellow]No enabled network adapters found[/yellow]")
            return
        
        if output_format == "json":
            output = [wmi_object_to_dict(adapter) for adapter in results]
            console.print(json.dumps(output, indent=2, default=str))
        else:
            for adapter in results:
                table = Table(title=f"Network Adapter: {adapter.Description}", show_header=False)
                table.add_column("Property", style="cyan bold")
                table.add_column("Value", style="white")
                
                table.add_row("MAC Address", adapter.MACAddress or "N/A")
                table.add_row("DHCP Enabled", str(adapter.DHCPEnabled))
                
                if adapter.IPAddress:
                    for i, ip in enumerate(adapter.IPAddress):
                        table.add_row(f"IP Address {i+1}", ip)
                
                if adapter.IPSubnet:
                    for i, subnet in enumerate(adapter.IPSubnet):
                        table.add_row(f"Subnet Mask {i+1}", subnet)
                
                if adapter.DefaultIPGateway:
                    for i, gw in enumerate(adapter.DefaultIPGateway):
                        table.add_row(f"Default Gateway {i+1}", gw)
                
                if adapter.DNSServerSearchOrder:
                    for i, dns in enumerate(adapter.DNSServerSearchOrder):
                        table.add_row(f"DNS Server {i+1}", dns)
                
                console.print(table)
                console.print()
            
    except Exception as e:
        console.print(f"[red]Error getting network info: {e}[/red]")
        raise typer.Exit(1)


@app.command()
def admin_check():
    """Check if running with administrator privileges."""
    if is_admin():
        console.print(Panel.fit(
            "[green]✓ Running with administrator privileges[/green]",
            title="Admin Status"
        ))
    else:
        console.print(Panel.fit(
            "[yellow]⚠ Not running with administrator privileges[/yellow]\n"
            "Some operations may require admin rights.",
            title="Admin Status"
        ))


@app.command()
def version():
    """Display version information."""
    from . import __version__
    console.print(f"[cyan]wmi-cli version {__version__}[/cyan]")


def _display_table(results: List, title: str = "Results"):
    """Helper to display WMI results as a table."""
    if not results:
        return
    
    # Get properties from first object
    sample = results[0]
    properties = []
    for prop in dir(sample):
        if prop.startswith('_'):
            continue
        try:
            value = getattr(sample, prop)
            if not callable(value):
                properties.append(prop)
        except:
            continue
    
    # Limit to reasonable number of columns
    if len(properties) > 10:
        properties = properties[:10]
    
    table = Table(title=title)
    for prop in properties:
        table.add_column(prop, style="cyan")
    
    for obj in results:
        row = []
        for prop in properties:
            try:
                value = getattr(obj, prop)
                row.append(str(value) if value is not None else "N/A")
            except:
                row.append("N/A")
        table.add_row(*row)
    
    console.print(table)


if __name__ == "__main__":
    app()
