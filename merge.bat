@echo off
REM Merge psychrometric chart changes to main repository

setlocal enabledelayedexpansion

set TOPIC_WORKTREE=d:\Climate-Zone-Finder\ClimateZoneFinder.worktrees\agents-add-psychrometric-chart-to-ppt
set MAIN_WORKTREE=d:\Climate-Zone-Finder\ClimateZoneFinder

echo.
echo ============================================================================
echo MERGING PSYCHROMETRIC CHART CHANGES TO MAIN REPOSITORY
echo ============================================================================
echo.

REM Step 1: Stage changes in topic worktree
echo [1/5] Staging changes in topic worktree...
cd /d "%TOPIC_WORKTREE%"
git add -A
if errorlevel 1 goto ERROR

REM Step 2: Create commit
echo [2/5] Creating commit...
git commit -m "feat: Add psychrometric chart to thermal comfort PPT report" -m "Add psychrometric chart visualization to thermal comfort analysis PowerPoint report with comfort zones and design strategy regions."
if errorlevel 1 (
    if not "%ERRORLEVEL%"=="1" goto ERROR
    echo (No changes to commit or error)
)

REM Step 3: Get current branch
echo [3/5] Getting branch name...
for /f "delims=" %%i in ('git branch --show-current') do set BRANCH=%%i
echo Current branch: !BRANCH!

REM Step 4: Merge to main worktree
echo [4/5] Merging to main repository...
git -C "%MAIN_WORKTREE%" merge !BRANCH!
if errorlevel 1 goto ERROR

REM Step 5: Verify
echo [5/5] Verifying merge...
git -C "%MAIN_WORKTREE%" log --oneline -3
echo.

echo ============================================================================
echo SUCCESS: Changes merged to main repository!
echo Location: %MAIN_WORKTREE%
echo ============================================================================
goto END

:ERROR
echo.
echo ERROR: Merge failed. Check output above.
exit /b 1

:END
