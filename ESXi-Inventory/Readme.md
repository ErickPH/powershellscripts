.SYNOPSIS
    Generates a comprehensive ESXi host inventory report with performance metrics, security checks, and multi-format outputs.
.DESCRIPTION
    Collects detailed ESXi host data including:
    - Host hardware, health stats, and performance metrics
    - Virtual Machines (config, snapshots, tools status, disk I/O)
    - Resource pools, clusters, and DRS settings
    - Network (vSwitches, VMkernel, physical NICs, firewall)
    - Storage (datastores, multipathing, I/O latency)
    - Security (users, roles, event logs)
.OUTPUTS
    JSON, HTML (styled), CSV, and optional email/Slack notifications.

### **How to Use the Script**

1. **Install PowerCLI** (if not already installed):
  
  powershell
  
    Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Force
  
2. **Run the Script**:
  
  # Basic run (all outputs):
.\ESXi_Inventory_Enhanced.ps1 -ESXiHost "192.168.1.100" -Username "root" -Password "yourpassword"

# With email notification:
.\ESXi_Inventory_Enhanced.ps1 -ESXiHost "192.168.1.100" -Username "root" -Password "yourpassword" -SendEmail -EmailTo "admin@yourdomain.com"

# With Slack alert:
.\ESXi_Inventory_Enhanced.ps1 -ESXiHost "192.168.1.100" -Username "root" -Password "yourpassword" -PostToSlack -SlackWebhook "https://hooks.slack.com/services/XXX"

3. **Output Files**:
  
  - **JSON**: `ESXi_Inventory.json` (structured data)
    
  - **HTML**: `ESXi_Inventory.html` (styled report)
    
  - **CSV**: Separate files for VMs, datastores, etc.
    

---

### **Key Features**

- **Complete Inventory**: Covers host hardware, VMs, datastores, and networks.
  
- **Flexible Outputs**: JSON (for APIs), HTML (readable report), CSV (Excel analysis).
  
- **Error Handling**: Validates PowerCLI installation and ESXi connectivity.
  
- **Styled HTML**: Professional-looking report with CSS.
  

---