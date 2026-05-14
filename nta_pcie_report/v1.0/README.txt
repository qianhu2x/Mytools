History:
260514_1200:
	Support to update the sheet "LTSSM_Stress".
	Note: The HSDES id has not been filled in yet.

Dependencies:
	Python 3.x
	pandas
	openpyxl

Usage:
	python nta_pcie_report.py <report.xlsx> <log_folder>
Arguments:
	ori_xlsx      Path to the original report Excel file to be updated.
	log_folder    Path to the folder containing test logs.
Options:
	-v, --version Display the script last update timestamp.

Example:
	python nta_pcie_report.py NTA_Report.xlsx ./logs/

Output:
	Updated report.xlsx with LTSSM_Stress sheet populated.
	Backup file created automatically: report_bak_YYMMDD_HHMMSS.xlsx

Notes:
	1. The script requires a sheet named LTSSM_Stress in the report file.
	2. Log filenames must contain GenX, pxpX, and portX patterns for matching.
        3. Files or folders with "deprecated" in their path are ignored.
	4. The original report file must not be open in another program.
	5. Debug logs (paths containing debug) are handled in a separate debug section.