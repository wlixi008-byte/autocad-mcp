import subprocess
import os
import sys
import shutil

def run_command(command, cwd=None):
    """Run a shell command and print output."""
    print(f"Running: {command}")
    try:
        # standard git/gh might output utf-8 even on windows, or use system codepage. 
        # Using errors='replace' is safest to prevent crashes.
        result = subprocess.run(
            command,
            cwd=cwd,
            shell=True,
            check=True,
            text=True,
            capture_output=True,
            encoding='utf-8', 
            errors='replace'
        )
        print(result.stdout)
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {command}")
        print(e.stderr)
        return False, e.stderr

def check_dependencies():
    """Check if git and gh are installed."""
    if not shutil.which("git"):
        print("Error: 'git' is not installed.")
        return False
    if not shutil.which("gh"):
        print("Error: 'gh' (GitHub CLI) is not installed.")
        return False
    return True

def main():
    print("--- GitHub Uploader Skill ---")
    
    if not check_dependencies():
        return

    # 1. Initialize Git
    if not os.path.exists(".git"):
        print("Initializing Git repository...")
        run_command("git init")
        # Create .gitignore if not exists
        if not os.path.exists(".gitignore"):
            with open(".gitignore", "w") as f:
                f.write("__pycache__/\n*.pyc\nvenv/\n.env\n.DS_Store\n")
    else:
        print("Git repository already initialized.")

    # 2. Check Auth Status
    print("Checking GitHub authentication...")
    success, output = run_command("gh auth status")
    if not success:
        print("You are not logged into GitHub CLI.")
        print("Please run 'gh auth login' in your terminal first, then re-run this script.")
        # In a real interactive Skill, we might try to trigger login, 
        # but for this script we fail gracefully.
        return

    # 3. Create Repo (if remote doesn't exist)
    # Check if remote 'origin' exists
    success, output = run_command("git remote get-url origin")
    if not success:
        print("Remote 'origin' not found. Creating new repository on GitHub...")
        repo_name = os.path.basename(os.getcwd())
        # Safe default: Private repo
        # Using --source=. to automaticall add remote and push?
        # gh repo create <name> --private --source=. --remote=origin
        cmd = f"gh repo create {repo_name} --private --source=. --remote=origin"
        
        # User provided username: wlixi008-byte
        # We can try to force it like: gh repo create wlixi008-byte/{repo_name}
        # But usually 'gh' uses the logged in user.
        
        success, _ = run_command(cmd)
        if not success:
            print("Failed to create repository. It might already exist.")
            # Try adding remote manually if creation failed but repo exists?
            # For now, let's stop/warn.
            print("Please check if the repository already exists on GitHub.")
    else:
        print(f"Remote 'origin' found: {output.strip()}")

    # 4. Add, Commit, Push
    print("Staging files...")
    run_command("git add .")
    
    # Check if there are changes to commit
    success, output = run_command("git status --porcelain")
    if output.strip():
        print("Committing changes...")
        run_command('git commit -m "Auto-update from GitHub Uploader Skill"')
        
        print("Pushing to GitHub...")
        # Push to main or master
        # Get current branch
        success, branch = run_command("git branch --show-current")
        branch = branch.strip()
        if not branch:
             branch = "main" # Fallback
             
        run_command(f"git push -u origin {branch}")
        print("Upload complete!")
    else:
        print("No changes to commit.")
        # Try push anyway just in case commits exist but aren't pushed
        run_command("git push")

if __name__ == "__main__":
    main()
