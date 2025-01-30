# Automated Processing of Employee Attendance Data

## Main Issue:
The **ZKTeco 6** device stores employee attendance data in a **read-only, encrypted Excel file**. This restriction prevents direct editing or processing of the data.  
The company needs to **extract this data in its standard format** and save it in a **new, editable Excel file**.

## Solution:
This software reads raw data from the **read-only file** and stores it in a **standard, editable, and report-friendly Excel file**.

## Features:
- **Reads attendance data**, including check-ins, check-outs, and absences, from the **read-only Excel file**  
- **Automatically processes the data** and converts it into a **new, editable Excel file**  
- **Categorizes information** by:
  - Date  
  - Day of the week  
  - Employee ID  
  - Check-in and check-out times  
- **Detects holidays and overtime** and records them in the output  
- **Stores each month's data** in a separate file  

## Outcome:
A **clear, well-organized, and editable Excel file**, suitable for **reporting and payroll calculations**.

# Program Workflow

1. **Receiving the raw attendance device Excel file (Read-Only)**  
2. **Automatically processing the data and converting it to a standard format**  
3. **Categorizing information by day, check-in and check-out times, and employee ID**  
4. **Generating a structured and editable output file**  
5. **Saving the processed data for reporting and payroll calculations**  

This entire process is completed in just a few seconds, with no need for manual processing.  

## Output Format  
Unlike the device’s original output, which stores data in multiple unorganized and non-editable sheets, this software saves the information in a **clean, readable, and editable Excel file**.

## Sample Output File:

| Row | Date       | Day       | Employee ID | Check-in | Check-out |
|----|-----------|-----------|-------------|---------|----------|
| 1  | 2024/01/21 | Saturday  | 1234        | 08:00   | 16:00    |
| 2  | 2024/01/22 | Sunday    | 1234        | 08:10   | 16:05    |
| 3  | 2024/01/23 | Monday    | 1234        | Absent  | Absent   |
| 4  | 2024/01/24 | Tuesday   | 5678        | 07:45   | 15:30    |
| 5  | 2024/01/25 | Wednesday | 5678        | 08:00   | 16:00    |

# Technologies Used in the Project

- **Python** for data processing  
- **Pandas, xlrd, xlwt, xlutils** for reading and writing Excel files  
- **JalaliDate and datetime** for managing Persian and Gregorian dates  
- **Tkinter and customtkinter** for designing the graphical user interface  

# How to Run This Program

## 1. Install Dependencies  
First, ensure that **Python 3** is installed on your system. Then, install the required libraries:  

```bash
pip install pandas xlrd xlwt xlutils persiantools customtkinter 
```

## 2. Run the Program  
After installing the libraries, simply execute the `main.py` file:  

```bash
python main.py 
```

## 3. Using the Software
- Select the Excel file exported from the attendance device.
- Choose the destination path for the processed file.
- Click the Process button.
- The final standardized and editable file will be generated and saved in the specified location.

# Graphical User Interface (GUI)


# Conclusion

- This software processes raw, non-editable attendance data and saves it in a **clear and standardized format**.  
- It **reduces human error**, improves **accuracy**, and **speeds up** employee data processing.  
- A **simple and practical solution** for HR and accounting teams dealing with attendance records.  

Optimize employee data management with this software!  

