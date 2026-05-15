#!/usr/bin/env python
"""Helper script to check git status and commit changes."""

import subprocess
import os

os.chdir(r'd:\Climate-Zone-Finder\ClimateZoneFinder.worktrees\agents-add-psychrometric-chart-to-ppt')

print("=== Recent commits (last 10) ===")
result = subprocess.run(['git', 'log', '--oneline', '-10'], capture_output=True, text=True)
print(result.stdout if result.stdout else "(none)")

print("\n=== Current status ===")
result = subprocess.run(['git', 'status', '--short'], capture_output=True, text=True)
print(result.stdout if result.stdout else "(clean)")

print("\n=== Staged changes ===")
result = subprocess.run(['git', 'diff', '--cached', '--stat'], capture_output=True, text=True)
print(result.stdout if result.stdout else "(none)")

print("\n=== Unstaged changes ===")
result = subprocess.run(['git', 'diff', '--stat'], capture_output=True, text=True)
print(result.stdout if result.stdout else "(none)")

# Stage all changes
print("\n=== Staging all changes ===")
result = subprocess.run(['git', 'add', '-A'], capture_output=True, text=True)
print("Staged")

# Get the diff for commit message
print("\n=== Changes to be committed ===")
result = subprocess.run(['git', 'diff', '--cached', '--stat'], capture_output=True, text=True)
print(result.stdout)
