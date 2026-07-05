# -*- coding: utf-8 -*-
import os
import sys
import time

try:
    import win32com.client
except ImportError:
    print("[-] Please install pywin32: pip install pywin32")
    sys.exit(1)

def monitor_processes():
    print("[*] Interrogating WMI for process creation events...")
    print("[*] Monitoring started. Press Ctrl+C to stop.\n")
    print(f"{'Time':<10} | {'User':<20} | {'PID':<8} | {'Process Name':<25} | {'Command Line'}")
    print("-" * 100)
    
    try:
        obj_wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        obj_wbem_services = obj_wmi_service.ConnectServer(".", "root\cimv2")
        
        query = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
        
        watcher = obj_wbem_services.ExecNotificationQuery(query)
        
        while True:
            try:
                event_process = watcher.NextEvent()
                process = event_process.TargetInstance
                
                proc_name = process.Name
                proc_id = process.ProcessId
                cmd_line = process.CommandLine if process.CommandLine else "N/A"
                
                try:
                    owner_info = process.GetOwner()
                    if owner_info[0] == 0:
                        user_owner = f"{owner_info[2]}\\{owner_info[1]}"
                    else:
                        user_owner = "Unknown"
                except Exception:
                    user_owner = "Access Denied"
                
                current_time = time.strftime("%H:%M:%S", time.localtime())
                print(f"{current_time:<10} | {user_owner:<20} | {proc_id:<8} | {proc_name:<25} | {cmd_line}")
                
            except KeyboardInterrupt:
                print("\n[*] Stopping process monitor. Goodbye!")
                break
            except Exception as e:
                continue
                
    except Exception as main_err:
        print(f"[-] Critical WMI Error: {main_err}")

if __name__ == "__main__":
    monitor_processes()
