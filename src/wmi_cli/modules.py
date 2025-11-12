"""
Advanced WMI query modules for specific use cases.
"""
from typing import List, Dict, Any, Optional
from .wmi_wrapper import WMIWrapper, wmi_object_to_dict


class ServiceManager:
    """Manage Windows services via WMI."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def start_service(self, service_name: str) -> tuple:
        """Start a Windows service."""
        services = self.wrapper.get_services(Name=service_name)
        if not services:
            raise ValueError(f"Service '{service_name}' not found")
        
        service = services[0]
        return service.StartService()
    
    def stop_service(self, service_name: str) -> tuple:
        """Stop a Windows service."""
        services = self.wrapper.get_services(Name=service_name)
        if not services:
            raise ValueError(f"Service '{service_name}' not found")
        
        service = services[0]
        return service.StopService()
    
    def restart_service(self, service_name: str) -> tuple:
        """Restart a Windows service."""
        self.stop_service(service_name)
        return self.start_service(service_name)
    
    def get_service_status(self, service_name: str) -> Dict[str, Any]:
        """Get detailed status of a service."""
        services = self.wrapper.get_services(Name=service_name)
        if not services:
            raise ValueError(f"Service '{service_name}' not found")
        
        return wmi_object_to_dict(services[0])
    
    def get_stopped_auto_services(self) -> List[Any]:
        """Get services that are set to start automatically but are stopped."""
        return self.wrapper.get_services(StartMode="Auto", State="Stopped")


class ProcessManager:
    """Manage processes via WMI."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def terminate_process(self, process_id: int) -> tuple:
        """Terminate a process by ID."""
        processes = self.wrapper.get_processes(ProcessId=process_id)
        if not processes:
            raise ValueError(f"Process with ID {process_id} not found")
        
        process = processes[0]
        return process.Terminate()
    
    def get_process_by_name(self, name: str) -> List[Any]:
        """Get all processes by name."""
        return self.wrapper.get_processes(Name=name)
    
    def get_process_info(self, process_id: int) -> Dict[str, Any]:
        """Get detailed information about a process."""
        processes = self.wrapper.get_processes(ProcessId=process_id)
        if not processes:
            raise ValueError(f"Process with ID {process_id} not found")
        
        return wmi_object_to_dict(processes[0])
    
    def get_high_memory_processes(self, min_memory_mb: int = 100) -> List[Any]:
        """Get processes using more than specified memory."""
        all_processes = self.wrapper.get_processes()
        min_bytes = min_memory_mb * 1024 * 1024
        return [p for p in all_processes if p.WorkingSetSize and int(p.WorkingSetSize) > min_bytes]


class SystemMonitor:
    """Monitor system resources and performance."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def get_cpu_info(self) -> List[Dict[str, Any]]:
        """Get CPU information."""
        processors = self.wrapper.get_processor()
        return [wmi_object_to_dict(p) for p in processors]
    
    def get_memory_info(self) -> Dict[str, Any]:
        """Get memory information."""
        cs = self.wrapper.get_computer_system()
        os_info = self.wrapper.get_operating_system()
        
        total_memory = int(cs.TotalPhysicalMemory)
        free_memory = int(os_info.FreePhysicalMemory) * 1024  # Convert KB to bytes
        used_memory = total_memory - free_memory
        
        return {
            "total_bytes": total_memory,
            "used_bytes": used_memory,
            "free_bytes": free_memory,
            "used_percentage": (used_memory / total_memory) * 100,
        }
    
    def get_disk_usage(self) -> List[Dict[str, Any]]:
        """Get disk usage for all local drives."""
        disks = self.wrapper.get_logical_disks(DriveType=3)  # Local disks only
        
        result = []
        for disk in disks:
            if disk.Size:
                size = int(disk.Size)
                free = int(disk.FreeSpace) if disk.FreeSpace else 0
                used = size - free
                
                result.append({
                    "device_id": disk.DeviceID,
                    "label": disk.VolumeName or "",
                    "file_system": disk.FileSystem,
                    "size_bytes": size,
                    "used_bytes": used,
                    "free_bytes": free,
                    "used_percentage": (used / size) * 100,
                })
        
        return result
    
    def get_uptime(self) -> Dict[str, Any]:
        """Get system uptime."""
        os_info = self.wrapper.get_operating_system()
        
        # LastBootUpTime is in WMI datetime format: 20231106120000.000000+000
        from datetime import datetime
        boot_time_str = os_info.LastBootUpTime.split('.')[0]
        boot_time = datetime.strptime(boot_time_str, "%Y%m%d%H%M%S")
        now = datetime.now()
        uptime = now - boot_time
        
        return {
            "boot_time": boot_time.isoformat(),
            "last_boot_time": boot_time.isoformat(),  # Added for compatibility with agent
            "current_time": now.isoformat(),
            "uptime_seconds": uptime.total_seconds(),
            "uptime_days": uptime.days,
            "uptime_hours": uptime.seconds // 3600,
            "uptime_minutes": (uptime.seconds % 3600) // 60,
        }


class NetworkManager:
    """Manage network configuration via WMI."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def get_active_adapters(self) -> List[Dict[str, Any]]:
        """Get all active network adapters."""
        adapters = self.wrapper.get_network_adapters(IPEnabled=True)
        return [wmi_object_to_dict(a) for a in adapters]
    
    def get_adapter_by_description(self, description: str) -> Optional[Dict[str, Any]]:
        """Get network adapter by description."""
        adapters = self.wrapper.get_network_adapters(Description=description, IPEnabled=True)
        return wmi_object_to_dict(adapters[0]) if adapters else None
    
    def get_network_statistics(self) -> List[Dict[str, Any]]:
        """Get network statistics for all adapters."""
        results = self.wrapper.query(
            "SELECT * FROM Win32_PerfRawData_Tcpip_NetworkInterface"
        )
        return [wmi_object_to_dict(r) for r in results]


class EventLogReader:
    """Read Windows Event Logs via WMI."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def get_event_logs(self) -> List[Dict[str, Any]]:
        """Get list of available event logs."""
        logs = self.wrapper.query("SELECT * FROM Win32_NTEventlogFile")
        return [wmi_object_to_dict(log) for log in logs]
    
    def get_recent_events(
        self,
        log_name: str = "System",
        event_type: Optional[int] = None,
        limit: int = 100
    ) -> List[Dict[str, Any]]:
        """
        Get recent events from a log.
        
        Args:
            log_name: Name of the log (System, Application, Security)
            event_type: Filter by type (1=Error, 2=Warning, 3=Information)
            limit: Maximum number of events to return
        """
        query = f"SELECT * FROM Win32_NTLogEvent WHERE Logfile = '{log_name}'"
        if event_type:
            query += f" AND Type = {event_type}"
        
        events = self.wrapper.query(query)
        # Limit results
        events = events[:limit]
        
        return [wmi_object_to_dict(e) for e in events]


class HardwareInfo:
    """Get hardware information via WMI."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def get_motherboard_info(self) -> Dict[str, Any]:
        """Get motherboard information."""
        boards = self.wrapper.get_class("Win32_BaseBoard")
        return wmi_object_to_dict(boards[0]) if boards else {}
    
    def get_video_controllers(self) -> List[Dict[str, Any]]:
        """Get video controller (GPU) information."""
        controllers = self.wrapper.get_class("Win32_VideoController")
        return [wmi_object_to_dict(c) for c in controllers]
    
    def get_sound_devices(self) -> List[Dict[str, Any]]:
        """Get sound device information."""
        devices = self.wrapper.get_class("Win32_SoundDevice")
        return [wmi_object_to_dict(d) for d in devices]
    
    def get_usb_controllers(self) -> List[Dict[str, Any]]:
        """Get USB controller information."""
        controllers = self.wrapper.get_class("Win32_USBController")
        return [wmi_object_to_dict(c) for c in controllers]
    
    def get_printers(self) -> List[Dict[str, Any]]:
        """Get printer information."""
        printers = self.wrapper.get_class("Win32_Printer")
        return [wmi_object_to_dict(p) for p in printers]
    
    def get_battery_status(self) -> Optional[Dict[str, Any]]:
        """Get battery status (for laptops)."""
        batteries = self.wrapper.get_class("Win32_Battery")
        return wmi_object_to_dict(batteries[0]) if batteries else None


class SecurityManager:
    """Security-related WMI queries."""
    
    def __init__(self, computer: str = "."):
        self.wrapper = WMIWrapper(computer=computer)
    
    def get_user_accounts(self) -> List[Dict[str, Any]]:
        """Get local user accounts."""
        accounts = self.wrapper.get_class("Win32_UserAccount", LocalAccount=True)
        return [wmi_object_to_dict(a) for a in accounts]
    
    def get_groups(self) -> List[Dict[str, Any]]:
        """Get local groups."""
        groups = self.wrapper.get_class("Win32_Group", LocalAccount=True)
        return [wmi_object_to_dict(g) for g in groups]
    
    def get_logged_on_users(self) -> List[Dict[str, Any]]:
        """Get currently logged-on users."""
        users = self.wrapper.get_class("Win32_LoggedOnUser")
        return [wmi_object_to_dict(u) for u in users]
    
    def get_shares(self) -> List[Dict[str, Any]]:
        """Get network shares."""
        shares = self.wrapper.get_class("Win32_Share")
        return [wmi_object_to_dict(s) for s in shares]
    
    def get_startup_programs(self) -> List[Dict[str, Any]]:
        """Get programs that run at startup."""
        startup = self.wrapper.get_class("Win32_StartupCommand")
        return [wmi_object_to_dict(s) for s in startup]
