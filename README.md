# LicenseCheck

## Overview
LicenseCheck is a Python script that remotely checks the Windows edition and license status of one or more hosts (by hostname or IP address) using PsExec and Windows' built-in `slmgr.vbs` tool. Results are saved to a CSV file, and debug information is logged for troubleshooting.

## Features
- Checks Windows license status for multiple hosts concurrently
- Uses PsExec for remote execution, or runs locally for localhost
- Input validation and hostname/IP resolution
- Results saved to `results.csv`
- Debug logs saved to `debug.log`
- Command-line options for timeout, appending results, and custom PsExec path

## Logic
1. **Input Handling**: Reads hostnames/IPs from `host.txt` or prompts the user.
2. **Validation**: Validates input and resolves IPs to hostnames.
3. **Concurrency**: Uses a thread pool to process multiple hosts in parallel.
4. **License Check**:
   - Pings each host to check if it's online.
   - If online, runs `slmgr.vbs /dli` via PsExec (or directly for localhost).
   - Parses output for Windows edition and license status.
   - Handles errors, timeouts, and unreachable hosts.
5. **Results**: Writes results to `results.csv` (append or overwrite).
6. **Logging**: Logs debug info and errors to `debug.log`.

## Usage
### Prerequisites
- Python 3.x
- PsExec.exe in the same directory as the script (or specify with `--psexec-path`)
- Hosts listed in `host.txt` (one per line), or provide interactively

### Command Line Options
```
python license_check.py [--timeout SECONDS] [--append] [--psexec-path PATH]
```
- `--timeout SECONDS` : Timeout for each remote command (default: 90)
- `--append` : Append results to `results.csv` instead of overwriting
- `--psexec-path PATH` : Path to PsExec.exe (default: script directory)

### Example
```
python license_check.py --timeout 120 --append --psexec-path "C:\Tools\PsExec.exe"
```

### Output
- Results are saved in `results.csv` with columns: Hostname, Windows Edition, License Status
- Example:
  ```csv
  Hostname,Windows Edition,License Status
  WIN2PC,"Windows(R), Enterprise edition",Licensed
  ```
- Debug information is saved in `debug.log`

## Security Notes
- Only trusted users should run this script (requires admin rights for PsExec)
- Input is validated and commands are constructed safely
- See script comments for further security recommendations

## Troubleshooting
- Check `debug.log` for errors and command output
- Ensure PsExec.exe is present and accessible
- Verify network/firewall settings for remote hosts

## License
This script is provided as-is for internal use. See comments in the script for further details.

## Version Control
This project includes a `.gitignore` file that excludes the following file types from version control:
- `.exe`
- `.txt`
- `.csv`
