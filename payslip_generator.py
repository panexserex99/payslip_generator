import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables for email credentials
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")


# Check that credentials are set
if not EMAIL_USER or not EMAIL_PASSWORD:
    raise Exception("Email credentials not found in the .env file.")

# Read employee data from Excel
def read_employee_data(filename="employees.xlsx"):
    try:
        df = pd.read_excel(filename)
        # Validate if any required column has missing values
        required_columns = ["Employee ID", "Name", "Email", "Basic Salary", "Allowances", "Deductions"]
        if any(col not in df.columns for col in required_columns):
            raise Exception(f"One or more required columns are missing. Required columns are: {required_columns}")
        if df[required_columns].isnull().any().any():
            raise Exception("Missing data in one or more required fields.")
        # Calculate Net Salary: Basic Salary + Allowances - Deductions
        df["Net Salary"] = df["Basic Salary"] + df["Allowances"] - df["Deductions"]
        return df
    except Exception as e:
        print("Error reading employee data:", e)
        raise

# Generate a PDF payslip for a given employee row using fpdf
def generate_payslip(row):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        
        # Header
        pdf.cell(0, 10, "Monthly Payslip", ln=True, align="C")
        pdf.ln(10)
        
        # Employee Details
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 10, f"Employee ID: {row['Employee ID']}", ln=True)
        pdf.cell(0, 10, f"Name: {row['Name']}", ln=True)
        pdf.ln(5)
        
        # Salary Details
        pdf.cell(0, 10, f"Basic Salary: ${row['Basic Salary']}", ln=True)
        pdf.cell(0, 10, f"Allowances: ${row['Allowances']}", ln=True)
        pdf.cell(0, 10, f"Deductions: ${row['Deductions']}", ln=True)
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, f"Net Salary: ${row['Net Salary']}", ln=True)
        
        # Ensure the payslips folder exists
        if not os.path.exists("payslips"):
            os.makedirs("payslips")
        
        # Save the PDF to payslips/{EmployeeID}.pdf
        pdf_filename = os.path.join("payslips", f"{row['Employee ID']}.pdf")
        pdf.output(pdf_filename)
        return pdf_filename
    except Exception as e:
        print(f"Error generating payslip for {row['Employee ID']}: ", e)
        raise

# Send email with the PDF payslip attached using yagmail
def send_email(row, attachment):
    subject = "Your Payslip for This Month"
    body = f"Dear {row['Name']},\n\nPlease find your payslip attached.\n\nRegards,\nHR Team"
    try:
        with yagmail.SMTP(EMAIL_USER, EMAIL_PASSWORD) as yag:
            yag.send(
                to=row["Email"],
                subject=subject,
                contents=body,
                attachments=attachment
            )
            print(f"Email sent successfully to {row['Email']}.")
    except Exception as e:
        print(f"Error sending email to {row['Email']}: ", e)

# Main function to process payslips
def main():
    print("Reading employee data...")
    df = read_employee_data()
    
    for index, row in df.iterrows():
        try:
            # Generate PDF payslip
            pdf_file = generate_payslip(row)
            print(f"Payslip generated for {row['Name']} ({row['Employee ID']}).")
            
            # Send the payslip via email
            send_email(row, pdf_file)
        except Exception as e:
            print(f"An error occurred while processing {row['Employee ID']}: {e}")

if __name__ == "__main__":
    main()
