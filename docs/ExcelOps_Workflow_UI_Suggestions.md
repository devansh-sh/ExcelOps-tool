# ExcelOps Workflow UI Suggestions

## Goal
Make ExcelOps friendly for non-technical users by separating the application into two clear modes: Create Workflow and Run Workflow.

## Recommended Startup Screen
When ExcelOps opens, show a simple landing screen instead of immediately opening the full editor.

Options:
- Create Workflow: opens the full ExcelOps tool for building or editing workflows.
- Run Workflow: opens a simplified workflow runner for regular users.
- Manage Workflows: optional screen for renaming, deleting, or reviewing saved workflows.

## Create Workflow Mode
Create Workflow should show the current full ExcelOps editor.

This mode is for users who build workflows and should include:
- Load files
- Filters
- Sorting
- Columns and calculations
- Pivot tables
- VLOOKUP steps
- Preview
- Save workflow
- Export workbook

Recommendation: rename Presets to Workflows throughout the UI.

Suggested labels:
- Presets -> Workflows
- Save Preset -> Save Workflow
- Load Preset -> Open Workflow
- Run Preset Workflow -> Run Workflow

## Run Workflow Mode
Run Workflow should hide the full editor and show only what a non-technical user needs.

Suggested flow:
1. Select a saved workflow.
2. Select the main input file.
3. Select any supporting lookup files required by saved VLOOKUP steps.
4. Click Run Workflow.
5. Preview the completed output.
6. Export the workbook.

The user should not see filter, pivot, VLOOKUP, formula, or column editor settings in this mode.

## Workflow Summary Before Running
After a workflow is selected, show a plain-language summary.

Example:
Workflow: Client Surveillance Report

This workflow will:
- Apply filters
- Create a pivot table
- Run 2 VLOOKUP steps
- Apply calculations
- Generate multiple output sheets

Required files:
- Main file
- Lookup file for VLOOKUP 1
- Lookup file for VLOOKUP 2

## Supporting File Prompts
File prompts should explain what each file is used for.

Example:
Select supporting file for VLOOKUP 1
Adds: Incomeamount
Matches: Clientcode -> Row Labels

This is better than a generic prompt such as: Select lookup file.

## Preview After Run Workflow
Preview should be available after the process completes.

Recommended completed workflow screen:
- Output sheet selector
- Row count
- Column count
- Preview table
- Export Workbook button
- Run Another Workflow button
- Back Home button

The preview screen should show the final generated outputs only, not the full workflow editor.

## Suggested App Structure
Recommended frontend structure:

ExcelOpsApp
- HomeScreen
- WorkflowEditorScreen
- WorkflowRunnerScreen
- WorkflowPreviewScreen
- WorkflowEngine

WorkflowEngine should handle actual processing so the UI stays simple.

## Suggested File Structure
Recommended files as the app grows:

- main.py
- home_screen.py
- workflow_editor.py
- workflow_runner.py
- workflow_preview.py
- workflow_engine.py
- presets.py
- filters.py
- sorts.py
- columns_manager.py
- pivot.py
- vlookup_frame.py
- vlookup_helper.py

## Packaging Recommendation
Use PyInstaller to package ExcelOps as a desktop app.

For macOS:
pyinstaller --windowed --name ExcelOps --icon excelops.icns main.py

For Windows:
pyinstaller --windowed --onefile --name ExcelOps --icon excelops.ico main.py

## Best Implementation Plan
Phase 1: Rename Presets to Workflows and add the startup screen.

Phase 2: Build the simplified Workflow Runner.

Phase 3: Add the completed-output Preview screen.

Phase 4: Package the app with PyInstaller and test on target machines.

## Final Recommended User Flow
Open ExcelOps.

Choose Create Workflow or Run Workflow.

If Create Workflow is selected:
- Open the full ExcelOps editor.
- Build or edit the workflow.
- Save the workflow.

If Run Workflow is selected:
- Select workflow.
- Select required files.
- Run workflow.
- Preview results.
- Export workbook.
