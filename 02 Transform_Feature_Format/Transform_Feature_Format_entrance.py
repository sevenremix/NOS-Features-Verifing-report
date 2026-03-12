import subprocess
import sys

def run_script(script_name):
    print(f"\n>>> Running {script_name}...")
    try:
        # Using sys.executable to ensure we use the same python interpreter
        result = subprocess.run([sys.executable, script_name], check=True, text=True)
        print(f">>> {script_name} completed successfully.")
        return True
    except subprocess.CalledProcessError as e:
        print(f">>> Error: {script_name} failed with exit code {e.returncode}")
        return False
    except Exception as e:
        print(f">>> An unexpected error occurred while running {script_name}: {e}")
        return False

def main():
    # Note: Using Transform_simple_copy.py as it is the current active script in the workspace
    scripts = [
        "Transform_simple_copy.py",
        "check_consistency.py"
        #"Feature_Case_Merging.py"
    ]
    
    for script in scripts:
        success = run_script(script)
        if not success:
            print("\n!!! Pipeline halted due to error.")
            sys.exit(1)
            
    print("\n✅ All steps completed successfully.")

if __name__ == "__main__":
    main()
