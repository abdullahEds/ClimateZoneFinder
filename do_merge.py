#!/usr/bin/env python
"""Merge script for psychrometric chart changes."""

import subprocess
import sys
import os

# Paths
topic_worktree = r'd:\Climate-Zone-Finder\ClimateZoneFinder.worktrees\agents-add-psychrometric-chart-to-ppt'
main_worktree = r'd:\Climate-Zone-Finder\ClimateZoneFinder'

def run_cmd(cmd, cwd=None, show_output=True):
    """Run a command and return output."""
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=cwd, shell=True)
    if show_output:
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr, file=sys.stderr)
    return result

print("=" * 60)
print("MERGE HELPER: Psychrometric Chart Changes")
print("=" * 60)

# Step 1: Check status in topic worktree
print("\n1. Checking status in topic worktree...")
os.chdir(topic_worktree)
result = run_cmd('git status --short')

# Step 2: Stage changes
print("\n2. Staging all changes...")
run_cmd('git add -A', show_output=False)
print("✓ Changes staged")

# Step 3: Show what will be committed
print("\n3. Changes to be committed:")
run_cmd('git diff --cached --stat')

# Step 4: Create commit with proper message
print("\n4. Creating commit...")
commit_msg = "feat: Add psychrometric chart to thermal comfort PPT report"
commit_body = """Add psychrometric chart visualization to thermal comfort analysis PowerPoint report.

Changes:
- New plot_psychrometric_chart() function in thermal_comfort_ppt.py
- Psychrometric chart slide with comfort zones and design strategy regions
- Updated cover slide to list all sections including psychrometric chart
- Chart displays ASHRAE 55 comfort zone and strategy zones
- Includes documentation of changes in PSYCHROMETRIC_CHART_CHANGES.md

The psychrometric chart shows:
- Hourly climate data points colored by dry bulb temperature
- Constant relative humidity curves (10-90% RH)
- ASHRAE 55 static comfort zone
- Design strategy zones (natural ventilation, evaporative cooling, heating)"""

run_cmd(f'git commit -m "{commit_msg}" -m "{commit_body}"')

print("\n5. Confirming commit...")
run_cmd('git log --oneline -1')

# Step 6: Get current branch name
print("\n6. Getting branch information...")
result = run_cmd('git branch --show-current', show_output=False)
current_branch = result.stdout.strip()
print(f"Current branch: {current_branch}")

# Step 7: Merge to main worktree
print(f"\n7. Merging to main worktree ({main_worktree})...")
merge_cmd = f'git -C {main_worktree} merge {current_branch}'
result = run_cmd(merge_cmd)

if result.returncode == 0:
    print("✓ Merge successful!")
    
    # Verify merge
    print("\n8. Verifying merge...")
    run_cmd(f'git -C {main_worktree} status --porcelain')
    
    print("\n" + "=" * 60)
    print("✓ MERGE COMPLETE")
    print("=" * 60)
else:
    print("✗ Merge failed or conflicts detected")
    print(result.stdout)
    if result.stderr:
        print("Error:", result.stderr)
    sys.exit(1)
