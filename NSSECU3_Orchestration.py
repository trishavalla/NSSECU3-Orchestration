import subprocess
import os
import time
import csv

# SET PATHS TO TOOLS
EVTXECMD_PATH = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\Tools\EvtxeCmd\EvtxECmd.exe"
RECMD_PATH = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\Tools\RECmd\RECmd.exe"
APPCOMPATCACHEPARSER_PATH = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\Tools\AppCompatCacheParser\AppCompatCacheParser.exe"

# INPUT FILE DIRECTORIES
EVTX_INPUT_DIR = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\evtx"
REGISTRY_INPUT_DIR = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\registry_hives"
OUTPUT_DIR = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\Output"
SYSTEM_FILE_PATH = os.path.join(REGISTRY_INPUT_DIR, "SYSTEM")

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

# OUTPUT FILES
EVTX_OUTPUT = os.path.join(OUTPUT_DIR, "evtx_output.csv")
RECMD_OUTPUT = os.path.join(OUTPUT_DIR, "recmd_output.csv")
APPCOMPAT_CSV = os.path.join(OUTPUT_DIR, "appcompat_output.csv")
FINAL_CSV_OUTPUT = os.path.join(OUTPUT_DIR, "forensic_results.csv")

# Function to run a command with error handling
def run_command(command, output_file):
    try:
        with open(output_file, "w") as f:
            subprocess.run(command, stdout=f, stderr=f, text=True, check=True)
        print(f"[+] Successfully executed: {' '.join(command)}")
    except subprocess.CalledProcessError as e:
        print(f"[!] Error executing {' '.join(command)}: {e}")

# Run EvtxECmd
def run_evtxecmd():
    if os.path.exists(EVTXECMD_PATH):
        cmd = [EVTXECMD_PATH, "-d", EVTX_INPUT_DIR, "--csv", OUTPUT_DIR]
        run_command(cmd, EVTX_OUTPUT)
    else:
        print(f"[!] EvtxECmd.exe not found at {EVTXECMD_PATH}")

# Run RECmd
def run_recmd():
    if os.path.exists(RECMD_PATH):
        RULE_FILE = r"C:\Users\trish\OneDrive\Desktop\MP2_NSSECU3\Tools\RECmd\BatchExamples\BatchExampleServices.reb"
        cmd = [RECMD_PATH, "-d", REGISTRY_INPUT_DIR, "--csv", OUTPUT_DIR, "--bn", RULE_FILE]
        run_command(cmd, RECMD_OUTPUT)
    else:
        print(f"[!] RECmd.exe not found at {RECMD_PATH}")

# Run AppCompatCacheParser
def run_appcompatcacheparser():
    if os.path.exists(APPCOMPATCACHEPARSER_PATH):
        cmd = [APPCOMPATCACHEPARSER_PATH, "-f", SYSTEM_FILE_PATH, "--csv", OUTPUT_DIR, "--csvf", "appcompat_output.csv"]
        
        print("[+] Running AppCompatCacheParser...")
        process = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        process.wait()

        while not os.path.exists(APPCOMPAT_CSV):
            print("[!] Waiting for AppCompatCacheParser output file to be created...")
            time.sleep(1)

        print(f"[+] Successfully processed AppCompatCacheParser output: {APPCOMPAT_CSV}")

    else:
        print(f"[!] AppCompatCacheParser.exe not found at {APPCOMPATCACHEPARSER_PATH}")

# Run the tools
print("[+] Running EvtxECmd...")
run_evtxecmd()
print("[+] Running RECmd...")
run_recmd()
print("[+] Running AppCompatCacheParser...")
run_appcompatcacheparser()

# Merge CSV files 
data_files = {
    "EvtxECmd": EVTX_OUTPUT,
    "RECmd": RECMD_OUTPUT,
    "AppCompatCacheParser": APPCOMPAT_CSV
}

def merge_csv_files():
    """Merge all forensic tool outputs into a structured CSV file without losing any entries."""
    with open(FINAL_CSV_OUTPUT, "w", encoding="utf-8", newline="") as f_out:
        writer = csv.writer(f_out)

        # Headers
        writer.writerow(["Source", "Message/Value"])

        for source, file_path in data_files.items():
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f_in:
                    reader = csv.reader(f_in)
                    headers = next(reader, None) 
                    if headers:
                        writer.writerow([source]+ headers)

                    # Write each row with the corresponding source
                    for row in reader:
                        writer.writerow([source] + row)
            else:
                print(f"[!] Warning: {source} output file not found ({file_path})")

    print(f"[+] Merged CSV saved to: {FINAL_CSV_OUTPUT}")

# Call the function to merge
merge_csv_files()