import os
import subprocess
import csv
import concurrent.futures
import argparse
import logging
import socket

# Define constants
LOCAL_NAMES = ['localhost', '127.0.0.1', os.environ.get('COMPUTERNAME', '').lower()]

def remove_domain(hostname):
    """Strip domain from hostname (e.g., 'host.domain.com' -> 'host')"""
    return hostname.split('.')[0] if '.' in hostname else hostname

def process_host(hostname_or_ip, psexec_path, existing_data, script_dir, timeout):
    """Process a single host to retrieve Windows edition and license status."""
    hostname_or_ip = hostname_or_ip.strip()
    if not hostname_or_ip:
        logging.error('Empty hostname/IP provided.')
        return None

    try:
        hostname = socket.gethostbyaddr(hostname_or_ip)[0]
    except socket.herror:
        hostname = hostname_or_ip
        logging.warning(f'Could not resolve hostname for {hostname_or_ip}, using input.')

    hostname = remove_domain(hostname)
    hostname_key = hostname.upper()

    # Skip licensed systems already recorded
    if existing_data.get(hostname_key, {}).get("License Status") == "Licensed":
        return None

    # Quick ping test
    ping_result = subprocess.run(
        ["ping", "-n", "1", "-w", "1000", hostname],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    if ping_result.returncode != 0:
        return hostname_key, "Offline", "N/A", "Ping failed"

    try:
        if hostname.lower() in LOCAL_NAMES:
            command = ['cscript', '//nologo', 'C:\\Windows\\System32\\slmgr.vbs', '/dli']
        else:
            command = [psexec_path, '-accepteula', f'\\\\{hostname}', 'cscript', '//nologo', 'C:\\Windows\\System32\\slmgr.vbs', '/dli']

        logging.info(f"Running: {' '.join(command)}")

        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=True,
            cwd=script_dir
        )

        output = result.stdout + result.stderr
        edition = status = "Unknown"

        for line in output.splitlines():
            if "Name:" in line:
                edition = line.split(":", 1)[1].strip()
            elif "License Status:" in line:
                status = line.split(":", 1)[1].strip()

        return hostname_key, edition, status, f"Stdout:\n{result.stdout}\nStderr:\n{result.stderr}"

    except subprocess.CalledProcessError as e:
        return hostname_key, "Error", "Error", f"CalledProcessError: {e}\nStderr: {e.stderr}"
    except subprocess.TimeoutExpired:
        return hostname_key, "Timeout", "N/A", f"TimeoutExpired after {timeout} seconds"

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Check Windows license status on multiple hosts.')
    parser.add_argument('hostnames', nargs='*', help='Hostnames or IPs to process')
    parser.add_argument('--timeout', type=int, default=90, help='Timeout in seconds for subprocess commands')
    parser.add_argument('--append', action='store_true', help='Append to results.csv instead of overwriting')
    parser.add_argument('--psexec-path', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'PsExec.exe'), help='Path to PsExec.exe')
    args = parser.parse_args()

    logging.basicConfig(filename='debug.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Load hostnames from args, host.txt, or prompt
    if args.hostnames:
        hostnames = [h.strip() for h in args.hostnames if h.strip()]
    elif os.path.exists("host.txt"):
        with open("host.txt", "r") as f:
            hostnames = [line.strip() for line in f if line.strip()]
    else:
        user_input = input("Enter hostname or IP: ").strip()
        hostnames = [user_input or "localhost"]

    # Load existing results
    existing_data = {}
    if os.path.exists("results.csv"):
        with open("results.csv", "r", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                key = remove_domain(row['Hostname'].strip()).upper()
                existing_data[key] = row

    # Run tasks concurrently
    results = []
    debug_logs = []
    max_workers = min(32, os.cpu_count() + 4)

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(process_host, host, args.psexec_path, existing_data, script_dir, args.timeout)
            for host in hostnames
        ]
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            if result:
                hostname, edition, status, debug_info = result
                results.append({'Hostname': hostname, 'Windows Edition': edition, 'License Status': status})
                debug_logs.append(f"Host: {hostname}\n{debug_info}\n")

    # Write results
    mode = 'a' if args.append else 'w'
    file_exists = os.path.exists("results.csv")

    with open("results.csv", mode, newline="") as f:
        fieldnames = ['Hostname', 'Windows Edition', 'License Status']
        writer = csv.DictWriter(f, fieldnames=fieldnames, lineterminator="\n")
        if mode == 'w' or not file_exists:
            writer.writeheader()
        for row in results:
            writer.writerow(row)

    # Output debug info
    print("Processing complete. Results saved to results.csv.")
    print("\nDebug Information:")
    for log in debug_logs:
        print(log)
