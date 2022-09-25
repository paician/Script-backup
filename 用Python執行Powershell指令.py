import subprocess

def run_powershell_command(command):
    completed = subprocess.run(["powershell", "-Command", command], capture_output=True)
    return completed


get_logs_command = 'Get-WinEvent -LogName Microsoft-Windows-TerminalServices-LocalSessionManager/Operational  | Where { ($_.ID -eq "25" -or  $_.ID -eq "21") -and ($_.TimeCreated -gt [datetime]::Today.AddDays(-28))} |Select TimeCreated , Message | sort-Object -Property TimeCreated -Unique | Format-List'
result = run_powershell_command(get_logs_command)

for line in result.stdout.splitlines():

    print(line.decode("ansi"))
