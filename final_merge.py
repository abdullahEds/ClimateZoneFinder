#!/usr/bin/env python
"""Complete merge workflow: commit in topic branch and merge to main."""

import subprocess
import sys
import os

topic_worktree = r'd:\Climate-Zone-Finder\ClimateZoneFinder.worktrees\agents-add-psychrometric-chart-to-ppt'
main_worktree = r'd:\Climate-Zone-Finder\ClimateZoneFinder'

def run_cmd(cmd, cwd=None):
    """Run a command and return result."""
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=cwd, shell=False)
    return result

print("=" * 70)
print("MERGING PSYCHROMETRIC CHART CHANGES TO MAIN REPOSITORY")
print("=" * 70)

try:
    # Step 1: Change to topic worktree
    print("\n[1/6] Working in topic branch worktree...")
    os.chdir(topic_worktree)
    
    # Step 2: Add all changes
    print("[2/6] Staging all changes...")
    result = run_cmd(['git', 'add', '-A'], cwd=topic_worktree)
    if result.returncode != 0:
        print(f"Error staging: {result.stderr}")
        sys.exit(1)
    
    # Step 3: Create commit
    print("[3/6] Creating commit...")
    commit_msg = "feat: Add psychrometric chart to thermal comfort PPT report"
    commit_body = """Add psychrometric chart visualization to thermal comfort analysis PowerPoint report.

Changes:
- New plot_psychrometric_chart() function in thermal_comfort_ppt.py
- Psychrometric chart slide with comfort zones and design strategy regions
- Updated cover slide to list all sections including psychrometric chart
- Chart displays ASHRAE 55 comfort zone and strategy zones
- Includes documentation in PSYCHROMETRIC_CHART_CHANGES.md

The psychrometric chart shows hourly climate data colored by temperature,
constant relative humidity curves, and design strategy zones."""
    
    result = run_cmd(['git', 'commit', '-m', commit_msg, '-m', commit_body], cwd=topic_worktree)
    if result.returncode != 0:
        print(f"Error committing: {result.stderr}")
        sys.exit(1)
    print(result.stdout)
    
    # Step 4: Get branch name
    print("[4/6] Getting branch information...")
    result = run_cmd(['git', 'branch', '--show-current'], cwd=topic_worktree)
    branch_name = result.stdout.strip()
    print(f"Current branch: {branch_name}")
    
    # Step 5: Merge to main worktree
    print(f"[5/6] Merging '{branch_name}' into main repository...")
    result = run_cmd(['git', '-C', main_worktree, 'merge', branch_name])
    
    if result.returncode != 0:
        print(f"Merge error: {result.stderr}")
        print(result.stdout)
        # Try to get more info about conflicts
        result2 = run_cmd(['git', '-C', main_worktree, 'status'], cwd=main_worktree)
        print("Current status:", result2.stdout)
        sys.exit(1)
    
    print(result.stdout)
    
    # Step 6: Verify
    print("[6/6] Verifying merge...")
    result = run_cmd(['git', '-C', main_worktree, 'log', '--oneline', '-3'])
    print("Last 3 commits in main:")
    print(result.stdout)
    
    print("\n" + "=" * 70)
    print("✓ SUCCESS: Changes merged to main repository!")
    print(f"Location: {main_worktree}")
    print("=" * 70)
    
except Exception as e:
    print(f"\n✗ ERROR: {e}")
    sys.exit(1)
