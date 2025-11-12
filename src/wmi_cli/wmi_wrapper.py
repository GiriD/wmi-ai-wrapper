"""
Core WMI wrapper module for interacting with Windows Management Instrumentation.
"""
import ctypes
import sys
from typing import Any, Dict, List, Optional, Union
from contextlib import contextmanager

try:
    import wmi
    import pythoncom
except ImportError as e:
    print(f"Error: Required packages not installed. Please install wmi and pywin32: {e}")
    sys.exit(1)


class WMIWrapper:
    """Main wrapper class for WMI operations."""
    
    _com_initialized = False
    
    def __init__(self, computer: str = ".", namespace: str = "root\\cimv2"):
        """
        Initialize WMI connection.
        
        Args:
            computer: Computer name or '.' for local machine
            namespace: WMI namespace (default: root\\cimv2)
        """
        self.computer = computer
        self.namespace = namespace
        self._connection: Optional[wmi.WMI] = None
        
        # Initialize COM once per class (not per instance)
        if not WMIWrapper._com_initialized:
            try:
                pythoncom.CoInitialize()
                WMIWrapper._com_initialized = True
            except:
                pass  # COM might already be initialized
    
    def get_connection(self):
        """Get or create WMI connection."""
        if self._connection is None:
            self._connection = wmi.WMI(computer=self.computer, namespace=self.namespace)
        return self._connection
    
    def query(self, wql_query: str) -> List[Any]:
        """
        Execute a WQL query.
        
        Args:
            wql_query: WQL query string
            
        Returns:
            List of query results
        """
        conn = self.get_connection()
        return list(conn.query(wql_query))
    
    def get_class(self, class_name: str, **kwargs) -> List[Any]:
        """
        Get instances of a WMI class.
        
        Args:
            class_name: Name of the WMI class (e.g., 'Win32_Service')
            **kwargs: Filter parameters
            
        Returns:
            List of class instances
        """
        conn = self.get_connection()
        wmi_class = getattr(conn, class_name)
        return list(wmi_class(**kwargs))
    
    def get_services(self, **filters) -> List[Any]:
        """Get Windows services."""
        return self.get_class("Win32_Service", **filters)
    
    def get_processes(self, **filters) -> List[Any]:
        """Get running processes."""
        return self.get_class("Win32_Process", **filters)
    
    def get_operating_system(self) -> Any:
        """Get operating system information."""
        results = self.get_class("Win32_OperatingSystem")
        return results[0] if results else None
    
    def get_computer_system(self) -> Any:
        """Get computer system information."""
        results = self.get_class("Win32_ComputerSystem")
        return results[0] if results else None
    
    def get_logical_disks(self, **filters) -> List[Any]:
        """Get logical disk drives."""
        return self.get_class("Win32_LogicalDisk", **filters)
    
    def get_network_adapters(self, **filters) -> List[Any]:
        """Get network adapter configuration."""
        return self.get_class("Win32_NetworkAdapterConfiguration", **filters)
    
    def get_bios(self) -> Any:
        """Get BIOS information."""
        results = self.get_class("Win32_BIOS")
        return results[0] if results else None
    
    def get_processor(self) -> List[Any]:
        """Get processor information."""
        return self.get_class("Win32_Processor")
    
    def get_physical_memory(self) -> List[Any]:
        """Get physical memory information."""
        return self.get_class("Win32_PhysicalMemory")
    
    def list_classes(self) -> List[str]:
        """List all available WMI classes in the current namespace."""
        conn = self.get_connection()
        # Query meta_class to get all available classes
        classes = []
        try:
            for cls in conn.query("SELECT * FROM meta_class"):
                class_name = str(cls.Path_.Class)
                if class_name:
                    classes.append(class_name)
        except Exception:
            # Fallback: return a list of common WMI classes
            classes = [
                "Win32_Service", "Win32_Process", "Win32_OperatingSystem",
                "Win32_ComputerSystem", "Win32_LogicalDisk", "Win32_NetworkAdapter",
                "Win32_NetworkAdapterConfiguration", "Win32_Processor", "Win32_BIOS",
                "Win32_PhysicalMemory", "Win32_BaseBoard", "Win32_VideoController",
                "Win32_SoundDevice", "Win32_Printer", "Win32_Battery",
                "Win32_UserAccount", "Win32_Group", "Win32_Share",
                "Win32_StartupCommand", "Win32_USBController", "Win32_DiskDrive",
                "Win32_NTEventlogFile", "Win32_PerfRawData_Tcpip_NetworkInterface"
            ]
        return sorted(set(classes))
    
    def get_class_properties(self, class_name: str) -> List[str]:
        """
        Get properties of a WMI class.
        
        Args:
            class_name: Name of the WMI class
            
        Returns:
            List of property names
        """
        conn = self.get_connection()
        wmi_class = getattr(conn, class_name)
        instances = list(wmi_class())
        if instances:
            return [prop for prop in dir(instances[0]) if not prop.startswith('_')]
        return []
    
    def call_method(self, instance: Any, method_name: str, *args, **kwargs) -> Any:
        """
        Call a method on a WMI instance.
        
        Args:
            instance: WMI instance object
            method_name: Name of the method to call
            *args, **kwargs: Method arguments
            
        Returns:
            Method result
        """
        method = getattr(instance, method_name)
        return method(*args, **kwargs)


def is_admin() -> bool:
    """
    Check if the script is running with administrator privileges.
    
    Returns:
        True if running as admin, False otherwise
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def require_admin(func):
    """Decorator to require admin privileges for a function."""
    def wrapper(*args, **kwargs):
        if not is_admin():
            raise PermissionError(
                "This operation requires administrator privileges. "
                "Please run as administrator."
            )
        return func(*args, **kwargs)
    return wrapper


def format_bytes(bytes_value: int) -> str:
    """
    Format bytes to human-readable format.
    
    Args:
        bytes_value: Size in bytes
        
    Returns:
        Formatted string (e.g., "1.5 GB")
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes_value < 1024.0:
            return f"{bytes_value:.2f} {unit}"
        bytes_value /= 1024.0
    return f"{bytes_value:.2f} PB"


def wmi_object_to_dict(wmi_object: Any, properties: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Convert WMI object to dictionary.
    
    Args:
        wmi_object: WMI object instance
        properties: List of properties to include (None for all)
        
    Returns:
        Dictionary representation of the object
    """
    if properties is None:
        properties = [prop for prop in dir(wmi_object) if not prop.startswith('_')]
    
    result = {}
    for prop in properties:
        try:
            value = getattr(wmi_object, prop)
            # Skip methods and special attributes
            if callable(value):
                continue
            result[prop] = value
        except Exception:
            result[prop] = None
    
    return result
