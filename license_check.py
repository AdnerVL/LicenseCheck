import os
import subprocess
import csv
import concurrent.futures
import threading
import argparse
import logging
import socket

def process_host(hostname_or_ip, lock, psexec_path, existing_data, script_dir, timeout):
    """Process a single host to retrieve Windows edition and license status."""
    # Input validation
    if not hostname_or_ip or not isinstance(hostname_or_ip, str):
        logging.error(f'Invalid input: {hostname_or_ip}')
        return None
    
    # Resolve IP to hostname if provided, otherwise use the input as hostname
    try:
        hostname = socket.gethostbyaddr(hostname_or_ip)[0]
    except socket.herror:
        hostname = hostname_or_ip  # Use input as hostname if resolution fails
        logging.warning(f'Could not resolve hostname for {hostname_or_ip}, using input as hostname')
    
    # Skip if host is already licensed in existing data
    if hostname.upper() in existing_data and existing_data[hostname.upper()][1] == "Licensed":
        return None
    
    # Check if host is reachable
    ping_result = subprocess.run(
        ["ping", "-n", "1", hostname],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    if ping_result.returncode == 0:
        try:
            local_names = ['localhost', '127.0.0.1', os.environ.get('COMPUTERNAME', '').lower()]
            if hostname.lower() in local_names:
                # Run directly for localhost
                command = ['cscript', '//nologo', 'C:\\Windows\\System32\\slmgr.vbs', '/dli']
            else:
                # Use PsExec for remote hosts
                command = [psexec_path, '-accepteula', f'\\\\{hostname}', 'cscript', '//nologo', 'C:\\Windows\\System32\\slmgr.vbs', '/dli']
            logging.info(f"Running command for {hostname}: {' '.join(command)}")
            
            # Execute the command
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=True,
                cwd=script_dir
            )
            # Parse output
            output = result.stdout + result.stderr
            edition = None
            status = None
            for line in output.splitlines():
                if "Name:" in line:
                    edition = line.split(":", 1)[1].strip()
                elif "License Status:" in line:
                    status = line.split(":", 1)[1].strip()
            edition = edition if edition else "Unknown"
            status = status if status else "Unknown"
            debug_info = f"Stdout:\n{result.stdout}\nStderr:\n{result.stderr}"
        except subprocess.CalledProcessError as e:
            edition, status = "Error", "Error"
            debug_info = f"CalledProcessError: {e}\nStderr: {e.stderr}"
        except subprocess.TimeoutExpired:
            edition, status = "Timeout", "N/A"
            debug_info = f"TimeoutExpired after {timeout} seconds"
    else:
        edition, status = "Offline", "N/A"
        debug_info = "Ping failed"
    return hostname.upper(), edition, status, debug_info

if __name__ == "__main__":
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Check Windows license status on multiple hosts.')
    parser.add_argument('--timeout', type=int, default=90, help='Timeout in seconds for subprocess commands')
    parser.add_argument('--append', action='store_true', help='Append to results.csv instead of overwriting')
    parser.add_argument('--psexec-path', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'PsExec.exe'), help='Path to PsExec.exe')
    args = parser.parse_args()
    
    # Set up logging
    logging.basicConfig(filename='debug.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # Set up paths and load hostnames
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.exists("host.txt"):
        with open("host.txt", "r") as f:
            hostnames = [line.strip() for line in f if line.strip()]
    else:
        hostname = input("Enter hostname or IP: ").strip()
        hostnames = [hostname if hostname else "localhost"]
    
    # Load existing data from CSV
    existing_data = {}
    if os.path.exists("results.csv"):
        with open("results.csv", "r", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # Skip header
            for row in reader:
                if len(row) >= 3 and row[0].strip():
                    existing_data[row[0].upper()] = row
    
    # Process hosts concurrently
    results = []
    debug_logs = []
    lock = threading.Lock()
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(process_host, hostname, lock, args.psexec_path, existing_data, script_dir, args.timeout) for hostname in hostnames]
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            if result:
                hostname, edition, status, debug_info = result
                results.append((hostname, edition, status))
                debug_logs.append(f"Host: {hostname}\n{debug_info}\n")
    
    # Write results to CSV
    with lock:
        mode = 'a' if args.append else 'w'
        with open("results.csv", mode, newline="") as f:
            writer = csv.writer(f, lineterminator="\n")
            if mode == 'w' or not os.path.exists('results.csv'):
                writer.writerow(['Hostname', 'Windows Edition', 'License Status'])
            for hostname, edition, status in results:
                if hostname.strip():
                    writer.writerow([hostname, edition, status])
    
    # Output results and debug information
    print("Processing complete. Results saved to results.csv.")
    print("\nDebug Information:")
    for log in debug_logs:
        print(log)