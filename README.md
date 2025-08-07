# NLExplorer Excel Joiner

## Overview
NLExplorer Excel Joiner is a desktop tool for reformatting and summarizing data exported from Neurolucida Explorer's Excel analysis output files. It provides a user-friendly interface to quickly generate summary tables from complex Excel sheets, making downstream analysis easier.

## Features
- **Marker Count Summary**: Aggregates marker counts by type and name across all sheets in an Excel file.
- **Dendrite Trees Summary**: Summarizes dendrite tree metrics (length, surface, volume) and provides both per-tree and per-sheet statistics.
- Simple graphical interface (PyQt6) for selecting analysis type and input files.

## Requirements
- Python 3.8+
- [pandas](https://pandas.pydata.org/)
- [PyQt6](https://pypi.org/project/PyQt6/)

## Installation
1. Clone this repository:
	```sh
	git clone https://github.com/AEWorthy/NLExplorer-Excel-Joiner.git
	cd NLExplorer-Excel-Joiner
	```
2. (Recommended) Create and activate a virtual environment or conda environment.
3. Install dependencies:
	```sh
	pip install -r requirements.txt
	```

## Usage
1. Run the application:
	```sh
	python main.py
	```
2. A window will appear with buttons for each analysis type.
3. Click a button (e.g., "Marker Count Summary" or "Dendrite Trees Summary").
4. Select your Excel file when prompted.
5. The tool will process the file and save a new summary Excel file in the same directory as your input.

## Input File Structure
Each analysis expects specific columns in each sheet of your Excel file:

### Marker Count Summary
- Each sheet must contain columns: `Type`, `Name`, `Qty of Markers`

### Dendrite Trees Summary
- Each sheet must contain columns: `Tree`, `Length Total(µm)`, `Surface Total(µm²)`, `Volume Total(µm³)`

If these column names are missing or formatted incorrectly, the tool will warn you and skip those sheets.

## Output
- The output is an Excel file named like `YourFile_Summary_Output.xlsx`.
- For Marker Count Summary: a table with marker counts by type and name, across all sheets.
- For Dendrite Trees Summary: two sheets—one with per-tree details, one with per-sheet summary statistics.

## Troubleshooting
- **No file selected**: You must select an Excel file when prompted.
- **Missing columns**: Ensure your Excel sheets have the required columns (see above).
- **No output generated**: If all sheets are missing required columns, no summary will be created.

## Contact
For questions or suggestions, open an issue on GitHub or contact the repository owner.
GitHub Repo: https://github.com/AEWorthy/NLExplorer-Excel-Joiner

