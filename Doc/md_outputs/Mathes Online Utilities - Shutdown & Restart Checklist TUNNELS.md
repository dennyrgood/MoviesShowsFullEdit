# Mathes Online Utilities - Shutdown & Restart Checklist TUNNELS

Extracted from PDF: Mathes Online Utilities - Shutdown & Restart Checklist TUNNELS.pdf

---

Mathes Online Utilities - Shutdown & Restart Checklist
SHUTDOWN PROCEDURE
Step 1: Stop All Services (PowerShell as Admin)
powershell

# Stop the tasks
Stop-ScheduledTask -TaskName "Cloudflared Tunnel"
Stop-ScheduledTask -TaskName "Flask Excel Backend"
Stop-ScheduledTask -TaskName "Flask Full Edit Backend"
# Disable auto-restart on reboot
Disable-ScheduledTask -TaskName "Cloudflared Tunnel"
Disable-ScheduledTask -TaskName "Flask Excel Backend"
Disable-ScheduledTask -TaskName "Flask Full Edit Backend"

Step 2: Update Landing Page to Show Offline Status
Edit D:\excel-web-interface\index.html in the CONFIG section:
Find the Movies & Shows card and change:
javascript

{
icon: "

",

title: "Movies & Shows",
description: "Full CRUD editor... (Currently offline)", // Add this
buttons: [
{
text: "Full Editor",
url: "...",
style: "disabled" // Change from "primary" to "disabled"
},
{
text: "Quick Add",
url: "...",
style: "disabled" // Change from "secondary" to "disabled"
}
]
}

Then push the update:

powershell

cd D:\excel-web-interface
git add index.html
git commit -m "Mark Movies & Shows as offline"
git push origin main

Wait 1-2 minutes for GitHub Pages to deploy.

Step 3: Verify Everything is Stopped
powershell

# Check ports are no longer listening
netstat -ano | findstr :5000
netstat -ano | findstr :5001
# Should return nothing
# Check cloudflared isn't running
Get-Process -Name cloudflared -ErrorAction SilentlyContinue
# Should return nothing
# Check Flask apps aren't running
Get-Process -Name python -ErrorAction SilentlyContinue | Where-Object {$_.Path -like "*Python313*"}
# Should return nothing related to your backends

Step 4: Test Shutdown Complete
Open browser and test:
•

https://api.ldmathes.cc/api/health → Should timeout/fail

•

https://api-edit.ldmathes.cc/api/health → Should timeout/fail

•

https://dennyrgood.github.io/ → Still loads (static, safe)

Status: All backend services are now offline. Your machine is no longer exposing anything to the internet.

Optional: Disable Cloudflare Tunnel Routes (Web Dashboard)
If you want to prevent any access attempts through Cloudflare:

1. Go to Cloudflare Zero Trust Dashboard: https://one.dash.cloudflare.com/
2. Navigate to Networks → Tunnels
3. Click on excel-backend tunnel
4. Go to Public Hostname tab
5. You'll see two routes:
• api.ldmathes.cc → http://localhost:5000
• api-edit.ldmathes.cc → http://localhost:5001
6. Click the 3 dots (⋮) next to each route → Delete
7. Confirm deletion
Note: Deleting routes is optional. If the local services are stopped, the routes won't work anyway. Delete them if
you want to be extra cautious or if you're shutting down for an extended period.

RESTART PROCEDURE
Step 1: Re-enable Task Scheduler Tasks (PowerShell as Admin)
powershell

# Re-enable the tasks
Enable-ScheduledTask -TaskName "Cloudflared Tunnel"
Enable-ScheduledTask -TaskName "Flask Excel Backend"
Enable-ScheduledTask -TaskName "Flask Full Edit Backend"

Step 2: Start Services
Option A: Reboot (Recommended)
powershell

Restart-Computer

All services will auto-start on boot.
Option B: Manual Start (No Reboot)

powershell

# Start each task manually
Start-ScheduledTask -TaskName "Cloudflared Tunnel"
Start-ScheduledTask -TaskName "Flask Excel Backend"
Start-ScheduledTask -TaskName "Flask Full Edit Backend"
# Wait 10 seconds for services to initialize
Start-Sleep -Seconds 10

Step 2A: Re-add Cloudflare Tunnel Routes (If Deleted During Shutdown)
If you deleted the tunnel routes from the Cloudflare dashboard during shutdown, you need to re-add them:
Method 1: Via Cloudflare Zero Trust Dashboard (Recommended)
1. Go to Cloudflare Zero Trust Dashboard: https://one.dash.cloudflare.com/
2. Navigate to Networks → Tunnels
3. Click on excel-backend tunnel
4. Go to Public Hostname tab
5. Click Add a public hostname
For api.ldmathes.cc:
• Subdomain: api
• Domain: ldmathes.cc
• Path: (leave blank)
• Type: HTTP
• URL: localhost:5000
• Click Save hostname
For api-edit.ldmathes.cc:

• Click Add a public hostname again
• Subdomain: api-edit
• Domain: ldmathes.cc
• Path: (leave blank)
• Type: HTTP
• URL: localhost:5001
• Click Save hostname
Method 2: Via Command Line
powershell

# Add DNS routes
cloudflared tunnel route dns excel-backend api.ldmathes.cc
cloudflared tunnel route dns excel-backend api-edit.ldmathes.cc
# Restart tunnel to pick up changes
Stop-ScheduledTask -TaskName "Cloudflared Tunnel"
Start-Sleep -Seconds 5
Start-ScheduledTask -TaskName "Cloudflared Tunnel"

Note: Your local config.yml file should still have the ingress rules. If routes were only deleted from the
dashboard, they'll automatically be recreated when the tunnel restarts.

Step 3: Update Landing Page to Show Online Status
Edit D:\excel-web-interface\index.html in the CONFIG section:
Find the Movies & Shows card and change:

javascript

{
icon: "

",

title: "Movies & Shows",
description: "Full CRUD editor...", // Remove "(Currently offline)"
buttons: [
{
text: "Full Editor",
url: "...",
style: "primary" // Change from "disabled" to "primary"
},
{
text: "Quick Add",
url: "...",
style: "secondary" // Change from "disabled" to "secondary"
}
]
}

Then push the update:
powershell

cd D:\excel-web-interface
git add index.html
git commit -m "Mark Movies & Shows as online"
git push origin main

Wait 1-2 minutes for GitHub Pages to deploy.

Step 4: Verify Services are Running

powershell

# Check Task Scheduler status
Get-ScheduledTask | Where-Object {$_.TaskName -like "*Flask*" -or $_.TaskName -like "*Cloudflared*"}
# All should show "Ready" or "Running"
# Check ports are listening
netstat -ano | findstr :5000
netstat -ano | findstr :5001
# Should show LISTENING for both ports
# Check cloudflared is running
Get-Process -Name cloudflared -ErrorAction SilentlyContinue
# Should show cloudflared.exe running
# Check Flask apps are running
Get-Process -Name python -ErrorAction SilentlyContinue | Where-Object {$_.Path -like "*Python313*"}
# Should show python.exe processes

Step 5: Test Endpoints
Backend Health Checks:
powershell

# Test port 5000 backend
Invoke-WebRequest -Uri https://api.ldmathes.cc/api/health
# Should return: {"status":"ok"}
# Test port 5001 backend
Invoke-WebRequest -Uri https://api-edit.ldmathes.cc/api/health
# Should return: {"status":"ok"}

Browser Tests:
•

https://dennyrgood.github.io/ → Landing page loads, health check shows "online", Movies & Shows
card active

•

https://dennyrgood.github.io/excel-web-interface/ → Quick-add form works

•

https://dennyrgood.github.io/MoviesShowsFullEdit/ → Full editor loads rows

Step 6: End-to-End Test

1. Go to https://dennyrgood.github.io/excel-web-interface/
2. Submit a test entry (Code: 999.1, Title: "Test Restart")
3. Verify entry appears in Excel file: D:\OneDrive\MS\MoviesShows.xlsx
4. Verify backup was created with timestamp
Status: All services restored and operational.

TROUBLESHOOTING
Services Won't Start
Check Task Scheduler:
powershell

# View task details
Get-ScheduledTask -TaskName "Flask Excel Backend" | Select-Object *
Get-ScheduledTask -TaskName "Flask Full Edit Backend" | Select-Object *
Get-ScheduledTask -TaskName "Cloudflared Tunnel" | Select-Object *

Manually start to see errors:
powershell

# Test Flask backends
cd D:\OneDrive\MS\
python excel_backend.py
# Watch for errors
# In new PowerShell window:
python excel_backend_full_edit.py
# Watch for errors
# Test Cloudflare Tunnel
D:\OneDrive\MS\cloudflared.exe tunnel run excel-backend
# Watch for errors

Ports Already in Use

powershell

# Find what's using the port
netstat -ano | findstr :5000
# Note the PID (last column)
# Kill the process
taskkill /PID [PID] /F
# Restart the service
Start-ScheduledTask -TaskName "Flask Excel Backend"

Tunnel Not Connecting
powershell

# Check tunnel config
Get-Content C:\Users\[YOUR-USERNAME]\.cloudflared\config.yml
# Verify credentials exist
Test-Path C:\Users\[YOUR-USERNAME]\.cloudflared\*.json
# Restart tunnel
Stop-ScheduledTask -TaskName "Cloudflared Tunnel"
Start-Sleep -Seconds 5
Start-ScheduledTask -TaskName "Cloudflared Tunnel"

GitHub Pages Not Loading
• Check repo at: https://github.com/dennyrgood/excel-web-interface
• Verify GitHub Pages is enabled: Settings → Pages
• Wait 1-2 minutes for deployment
• Hard refresh browser: Ctrl + F5

QUICK REFERENCE
Current Configuration
Service

Port

Tunnel URL

Task Name

Quick-Add Backend

5000

api.ldmathes.cc

Flask Excel Backend

Full Editor Backend

5001

api-edit.ldmathes.cc

Flask Full Edit Backend

Cloudflare Tunnel

N/A

N/A

Cloudflared Tunnel

File Locations
Item

Location

Python

D:\Misc\Python313\python.exe

Backend Scripts

D:\OneDrive\MS\

Quick-Add

D:\OneDrive\MS\excel_backend.py

Full Editor

D:\OneDrive\MS\excel_backend_full_edit.py

Excel File

D:\OneDrive\MS\MoviesShows.xlsx

Cloudflared

D:\OneDrive\MS\cloudflared.exe

Tunnel Config

C:\Users\[USERNAME]\.cloudflared\config.yml

GitHub Repo

D:\excel-web-interface\

Key URLs
Purpose

URL

Landing Page

https://dennyrgood.github.io/

Quick-Add Form

https://dennyrgood.github.io/excel-web-interface/

Full Editor

https://dennyrgood.github.io/MoviesShowsFullEdit/

Port 5000 Health

https://api.ldmathes.cc/api/health

Port 5001 Health

https://api-edit.ldmathes.cc/api/health

OPTIONAL: COMPLETE REMOVAL
If you want to permanently remove the project (not just pause):

1. Delete Task Scheduler Tasks
powershell

Unregister-ScheduledTask -TaskName "Cloudflared Tunnel" -Confirm:$false
Unregister-ScheduledTask -TaskName "Flask Excel Backend" -Confirm:$false
Unregister-ScheduledTask -TaskName "Flask Full Edit Backend" -Confirm:$false

2. Delete Cloudflare Tunnel

powershell

# List tunnels
D:\OneDrive\MS\cloudflared.exe tunnel list
# Delete tunnel
D:\OneDrive\MS\cloudflared.exe tunnel delete excel-backend

Note: This permanently deletes the tunnel. If you restart later, you'll need to create a new tunnel with a new ID.

3. Remove DNS Records
Option A: Via Cloudflare Dashboard (Recommended)
1. Go to Cloudflare Zero Trust Dashboard: https://one.dash.cloudflare.com/
2. Navigate to Networks → Tunnels
3. Click on excel-backend tunnel
4. Go to Public Hostname tab
5. Delete both routes:
• api.ldmathes.cc
• api-edit.ldmathes.cc
6. Click the 3 dots (⋮) next to the tunnel name → Delete tunnel
Option B: Via Cloudflare DNS Dashboard
1. Go to Cloudflare Dashboard: https://dash.cloudflare.com/
2. Select your domain: ldmathes.cc
3. Go to DNS → Records
4. Find and delete CNAME records:
• api → [tunnel-id].cfargotunnel.com
• api-edit → [tunnel-id].cfargotunnel.com

4. Delete GitHub Pages (Optional)
• Go to https://github.com/dennyrgood/excel-web-interface/settings/pages
• Select "None" under Source
• Or delete entire repository

5. Clean Up Local Files (Optional)

powershell

# Remove backend scripts
Remove-Item D:\OneDrive\MS\excel_backend.py
Remove-Item D:\OneDrive\MS\excel_backend_full_edit.py
Remove-Item D:\OneDrive\MS\cloudflared.exe
# Remove tunnel credentials
Remove-Item -Recurse C:\Users\[USERNAME]\.cloudflared\
# Remove GitHub repo clone
Remove-Item -Recurse D:\excel-web-interface\

NOTES
• Pause vs. Remove: Disabling tasks = pause (can restart easily). Deleting tasks/tunnel = permanent
removal.
• Excel File Safety: Always remains local at D:\OneDrive\MS\MoviesShows.xlsx , never deleted in any
scenario.
• GitHub Pages: Stays live even when backends are stopped (forms just won't work).
• Domain: ldmathes.cc remains active and registered—only the subdomains stop working.
• Security: When shutdown, nothing is exposed to the internet. Your machine is secure.

LESSONS LEARNED
Great things accomplished in this project:
•

Set up Cloudflare Tunnel for secure external access

•

Built Flask REST APIs with full CRUD operations

•

Preserved Excel formulas and formatting programmatically

•

Created responsive web forms with GitHub Pages

•

Implemented multi-port backend architecture

•

Automated services with Windows Task Scheduler

•

Built configuration-driven, modular frontend

This was a successful experiment. You learned a ton about web architecture, Flask, Cloudflare, Excel
automation, and system administration. Well done!

