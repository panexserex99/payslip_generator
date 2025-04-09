#  Payslip Generator – Python Project

This Python project automates the creation and delivery of employee payslips. It reads employee details from an Excel file, calculates net salaries, generates PDF payslips, and sends them via email.

---

##  Features

- Reads employee data from `employees.xlsx`
- Calculates net salary using:
  ```
  Net Salary = Basic Salary + Allowances - Deductions
  ```
- Generates a professional PDF payslip for each employee
- Emails each payslip to the employee securely
- Handles missing data and email errors gracefully

---

##  Project Structure

```
payslip_generator/
│
├── payslip_generator.py      # Main Python script
├── employees.xlsx            # Excel file with employee details
├── .env                      # Stores email credentials securely
├── payslips/                 # Folder for generated PDF payslips
└── README.md                 # Project instructions
```

---

##  Requirements

Install required Python libraries:

```bash
pip install pandas fpdf yagmail python-dotenv openpyxl
```

---

## Setup Instructions

1. **Clone or download** the project folder.
2. Place your `employees.xlsx` file in the project folder with these columns:

   | Employee ID | Name | Email | Basic Salary | Allowances | Deductions |
   |-------------|------|-------|--------------|------------|-------------|

3. **Create a `.env` file** with your email credentials:

```
EMAIL_USER=your_email@gmail.com
EMAIL_PASSWORD=your_app_password
```

> **For Gmail:** Enable 2-Step Verification and create an [App Password](https://myaccount.google.com/apppasswords).

---

##  How to Run the Script

In your terminal or VS Code:

```bash
python payslip_generator.py
```

- PDF payslips will be generated in the `payslips/` folder.
- Each employee will receive an email with their payslip attached.

---

##  Email Format

- **Subject:** Your Payslip for This Month
- **Body:**
  ```
  Dear [Employee Name],

  Please find your payslip attached.

  Regards,  
  HR Team
  ```

---

##  Security Note

Never share your `.env` file or real credentials publicly. Add `.env` to `.gitignore` if using Git.

---

##  To Do

- [ ] Add a simple GUI (optional)
- [ ] Log failed emails to a file
- [ ] Add unit tests for salary calculation

---

## Contact

Made with 💻 by **Panashe Seremani**  
Uncommon.org – Empowering Through Tech 

---
