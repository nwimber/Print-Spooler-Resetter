# Print Spooler Resetter v0.1 4/6/23
# Requires pypsrp and requests-kerberos

from pypsrp.client import Client
from pypsrp.powershell import PowerShell, RunspacePool
from pypsrp.wsman import WSMan

# Connect to the remote Windows Server using PowerShell and Kerberos, currently testing without ssl
def connect_to_server(hostname, username, password):
    client = Client(
        hostname,
        auth="kerberos",
        ssl=False,
        port=5985,
        username=username,
        password=password,
        encryption="never",
        operation_timeout_sec=120,
    )
    return client


# Stop the print spooler service
def stop_print_spooler_service(client):
    output = client.execute_ps("Stop-Service -Name Spooler")
    print(output)

# Delete the contents of the print spooler directory
def delete_printer_directory_contents(client):
    ps_script = """
    $printerDirectory = "C:\Windows\System32\spool\PRINTERS"
    Get-ChildItem -Path $printerDirectory -Recurse | Remove-Item -Force
    """
    output = client.execute_ps(ps_script)
    print(output)

# Start the print spooler service
def restart_print_spooler_service(client):
    output = client.execute_ps("Start-Service -Name Spooler")
    print(output)

def main():
    # Prompt the user for required variables
    hostname = input("Please enter the server hostname: ")
    domain = input("Please enter your domain: ")
    username = input(f"Please enter your username for the {domain} domain: ")
    password = input("Please enter your password: ")

    # Format the username with the domain
    username = f"{domain}\\{username}"

    # Connect to the remote Windows Server
    client = connect_to_server(hostname, username, password)

    # Stop the print spooler service
    stop_print_spooler_service(client)

    # Delete the contents of the print spooler directory
    delete_printer_directory_contents(client)

    # Restart the print spooler service
    restart_print_spooler_service(client)

if __name__ == "__main__":
    main()
    
