# leader-process_monitor

# Windows WMI Process Monitor (`leader-process_monitor.py`)

A highly efficient, real-time Windows process monitoring tool written in Python 3 utilizing the native **Windows Management Instrumentation (WMI)** architecture. This tool is designed for security professionals, malware analysts, and penetration testers to gain deep visibility into process creation events, user ownership, and command-line execution strings.

---

## 🔬 Operational Use Cases

### 🔴 Red Teaming & Privilege Escalation
In a post-exploitation phase, escalating privileges from a low-privilege user to `NT AUTHORITY\SYSTEM` requires identifying misconfigured services, automated administrative tasks, or unquoted service paths. 
This tool hooks into the WMI subsystem to capture:
* **Hidden/Ephemeral Processes:** Detects fast-executing or short-lived binaries that bypass standard task managers.
* **Privileged Execution Traces:** Flags processes running under high-privilege context (`SYSTEM` / `Administrator`) that call external, user-writable scripts (`.vbs`, `.bat`, `.ps1`), paving the way for **Binary Hijacking** or **Task Scheduling exploits**.

### 🔵 Blue Teaming & Threat Hunting
For defensive operations, this script acts as a lightweight host-based monitoring mechanism to:
* Audit active command-line arguments for suspicious patterns (e.g., obfuscated PowerShell strings, unexpected `net user` creations, or unauthorized network callbacks).
* Track dynamic user execution behavior across the local machine.

---

## 🏛️ Core Architecture & Mechanics

The script interfaces directly with the Windows COM subsystem using `win32com.client` to execute a low-overhead notification query:

```sql
SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'


Dynamic WQL Event Polling: Instead of resource-heavy infinite polling loops, it leverages WMI's internal event subscription mechanism, checking every 1-second interval (WITHIN 1) for new Win32_Process instances.

Kernel-to-User Context Resolution: Dynamically invokes GetOwner() methods on intercepted process handles to extract security identifier (SID) boundaries and resolve domain/username mappings in real time.

🛠️ Prerequisites & Installation
This script is native to Windows environments and requires Python 3.x.

Clone the repository:

git clone [https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git)
cd YOUR_REPO_NAME

Install the necessary Python-to-Win32 bindings (pywin32):

pip install pywin32

🚀 Usage
⚠️ Security Note: To resolve user ownership (GetOwner) for highly privileged system processes (SYSTEM / Network Service), the script must be executed from an elevated command prompt (Run as Administrator).

Run the monitor using:

python leader-process_monitor.py

Expected Output Format:
The tool outputs a structured, easy-to-read live table tracking events instantly:

Time       | User                 | PID      | Process Name              | Command Line
----------------------------------------------------------------------------------------------------
14:32:05   | NT AUTHORITY\SYSTEM  | 4104     | svchost.exe               | C:\Windows\system32\svchost.exe -k netsvcs -p
14:32:12   | DESKTOP-LAB\pentester| 8940     | notepad.exe               | "C:\Windows\system32\notepad.exe" C:\Users\Public\notes.txt
14:32:18   | Access Denied        | 924      | DynamicPrivilegedProc.exe | N/A

14:32:18   | Access Denied        | 924      | DynamicPrivilegedProc.exe | N/A


🛑 Disclaimer
This tool is developed strictly for educational purposes, authorized security auditing, and defensive research. The author accepts no liability for any misuse or damage caused by this software.

