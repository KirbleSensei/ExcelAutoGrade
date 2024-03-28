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

### 1. Defining Expected Values and Whitelisted Formulas

Before running the script, define the expected values and whitelisted formulas in the script itself:

```python
# List of expected values (MUST BE IN THE SAME ORDER AS CELLS)
expected = (46, 47, 197)

# Whitelisted formulas (MUST BE IN THE SAME ORDER AS CELLS)
whitelist = ("=SUM(D2:D12)", "=SUM(E2:E12)", "=SUM(F2:F12)")
```
## Output
The script will process each Excel file in the specified folder, checking the defined cell range for compliance with the expected values and whitelisted formulas. It will print the grades for each student based on the assessment.
