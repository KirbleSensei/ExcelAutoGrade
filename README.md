# Excel Auto Grade

This Python script is designed to automate the grading process for Excel files. It allows users to define expected values and whitelisted formulas and then checks the specified range of cells in Excel files for compliance.

## Installation

1. Clone the repository:

    ```
    git clone https://github.com/your_username/ExcelAutoGrade.git
    ```

2. Navigate to the project directory:

    ```
    cd ExcelAutoGrade
    ```

3. Install the required libraries:

    ```
    pip install openpyxl
    pip install rarfile
    ```

## Usage

### File Structure

- `Warnings.txt`: Contains warnings for formulas used that are not in the whitelist.
- `Grades.txt`: Contains the grades assigned to each student based on the evaluation criteria.

### Notes

- The script assumes that the Excel files to be graded are located in a RAR archive (`Project01.rar`). Modify the `pathToZip` variable accordingly if your files are stored differently.
- Ensure that the expected values and whitelisted formulas are provided in the correct order as specified in the script.

```python
pathToZip = r"~\Project01.rar" # Path to initial archive file
Sheet = "Sheet1" # Sheet name to check
Cells = "D13:F13" # Range of Cells to check
# List of expected values (MUST BE IN THE SAME ORDER AS CELLS)
Expected = (46, 47, 197)
# Whitelisted formulas (MUST BE IN THE SAME ORDER AS CELLS)
Whitelist = ("=SUM(D2:D12)", "=SUM(E2:E12)", "=SUM(F2:F12)")

assertEqualsCells(pathToZip, Sheet, Cells, Expected, Whitelist)
```
## Output
The script will process each Excel file in the specified folder, checking the defined cell range for compliance with the expected values and whitelisted formulas. It will print the grades for each student based on the assessment.
