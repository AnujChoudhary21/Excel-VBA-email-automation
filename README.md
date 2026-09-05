# Excel-VBA-email-automation
Excel VBA Customer Email Automation
An AI-assisted Excel VBA automation project for streamlining
customer payment-reminder communication.
The workflow connects Excel, Microsoft Word, PDF generation, and
Outlook to transform customer data into personalized documents and
email communications with minimal manual effort.
> **Development note:** AI was used to generate and refine the VBA code
> based on the required workflow. The Excel integration, workbook
> configuration, buttons/Ribbon setup, testing, troubleshooting, and
> overall automation workflow were implemented and validated manually.
🚀 Project Overview
This project automates a three-stage business workflow:
``` text
Customer Data in Excel
        │
        ▼
1. Generate Personalized PDFs
        │
        ▼
2. Map PDF File Paths to Excel
        │
        ▼
3. Prepare & Send Personalized Emails
```
Instead of manually creating letters, locating individual PDF files, and
preparing emails one by one, the workflow uses Excel VBA to automate the
repetitive steps.
✨ Key Features
Generate personalized payment-reminder PDFs from Excel data
Replace dynamic fields such as `<<Customer Name>>` and
`<<Amount Pending>>`
Save generated PDFs using the corresponding Employee ID
Select the Word template through a file-selection dialog
Select the destination folder for generated PDFs
Search for matching PDF files using Employee IDs
Write the complete PDF file path back into Excel
Generate personalized customer emails
Attach the corresponding PDF automatically
Preview emails before sending
Execute automation through Excel buttons
Add frequently used macros to the Excel Ribbon/Quick Access
interface
Use input validation and basic error handling within the automation
workflow
🧩 Workflow
1. Generate Personalized PDFs
The Excel sheet contains customer information such as:
Column   Field
---
A        Employee ID
B        Customer Name
D        Amount Pending
The automation reads each row and replaces the corresponding fields in
the Word template:
``` text
<<Customer Name>>
<<Amount Pending>>
```
A separate PDF is then generated for each customer and saved using the
Employee ID.
Example:
``` text
1001.pdf
1002.pdf
1003.pdf
```
2. Map PDF File Paths
The automation can search a selected folder for PDF files matching
Employee IDs and write the corresponding full file path back into Excel.
Example:
``` text
Employee ID    PDF File Path
1001           C:\Demo\PDFs@1.pdf
1002           C:\Demo\PDFs@2.pdf
```
3. Personalized Email Automation
The email automation reads customer information from Excel and creates
an individualized email.
Example:
``` text
Hello Aarav Sharma,

Your Amount is Pending Rs.3500

Please check the Attachment File for more details.

Thanks & Regards,
Anuj Choudhary
```
The matching PDF is automatically attached.
The workflow includes a preview-before-send step, allowing the user
to review the generated message and attachment before sending.
📊 Input Data Structure
Column   Description
---
A        Employee ID
B        Customer Name
C        Email ID
D        Amount Pending
E        PDF File Path / Attachment
The project is designed to work with data containing headers.
📝 Word Template
The Word template uses placeholder fields that are replaced
automatically:
``` text
<<Customer Name>>
<<Amount Pending>>
```
This allows the same template to be reused for multiple customers.
A sanitized demonstration template is included in the `demo` folder.
🛠️ Technology Stack
Microsoft Excel
VBA (Visual Basic for Applications)
Microsoft Word
Microsoft Outlook
PDF
AI-assisted development
📁 Repository Structure
``` text
excel-vba-email-automation/
│
├── demo/
│   ├── CustomerData_Demo.xlsx
│   └── Letter_Template_Demo.docx
│
├── vba/
│   ├── GenerateCustomerPDFs.bas
│   ├── MapPDFPaths.bas
│   └── SendCustomerEmails.bas
│
├── screenshots/
│   ├── excel-customer-data.png
│   ├── generated-pdfs.png
│   ├── pdf-path-mapping.png
│   ├── email-preview.png
│   └── excel-ribbon-automation.png
│
└── README.md
```
📸 Screenshots
Customer Data
<img width="1066" height="54" alt="image" src="https://github.com/user-attachments/assets/3d12f3d0-183c-47ce-aeb6-82671995ef2f" />

Generated PDFs
<img width="884" height="345" alt="image" src="https://github.com/user-attachments/assets/3d92d0f3-476c-4f64-9689-6ac0ad148b21" />

PDF Path Mapping
<img width="1636" height="217" alt="image" src="https://github.com/user-attachments/assets/a0d5d3b5-642e-4c27-9fc4-ca349e4caad6" />

Email Preview
![Email Preview](screenshots/email-preview.png)
Excel Ribbon Automation
![Excel Ribbon Automation](screenshots/excel-ribbon-automation.png)
▶️ How to Use
Prepare the Excel data using the required headers.
Run the PDF-generation macro and select the Word template.
Select the destination folder for generated PDFs.
Run the PDF-path mapping macro and select the PDF folder.
Run the email automation.
Review the personalized email and attachment.
Send the email after confirmation.
🎯 Business Problem Solved
Manual customer communication can involve:
``` text
Find customer data
      ↓
Create individual letter
      ↓
Generate PDF
      ↓
Find attachment
      ↓
Prepare email
      ↓
Attach PDF
      ↓
Send email
```
This project consolidates these repetitive activities into:
``` text
Excel Data
    ↓
Automated PDF Generation
    ↓
Automated Attachment Mapping
    ↓
Personalized Email
    ↓
Preview
    ↓
Send
```
The goal is to reduce repetitive manual work, improve consistency, and
make the process easier to execute.
🤖 AI-Assisted Development
The development process involved:
Defining the required business workflow
Breaking the workflow into individual automation tasks
Using AI to generate and refine VBA solutions
Implementing the generated VBA in Excel
Configuring buttons and Excel interface controls
Testing the complete workflow
Troubleshooting implementation issues
Validating the final automation from an end-user perspective
This demonstrates requirements interpretation, workflow design,
AI-assisted development, Excel integration, testing, and automation
usability.
🔐 Data Privacy
The repository uses sanitized demonstration data.
Real customer names, email addresses, financial information,
confidential documents, and personal file paths should not be uploaded
to a public repository.
⚠️ Notes
The automation is intended for Microsoft Excel/VBA environments that
support macros.
Email automation depends on the Outlook setup and permissions on the
user's computer.
Macros may require appropriate Excel macro/security settings.
Always review generated emails before sending them in a real-world
workflow.
👤 Author
Anuj Choudhary
AI-Assisted Excel & VBA Automation Project
⭐ Project Purpose
A practical demonstration of how AI-assisted development can be 
combined with Excel VBA to automate a real-world administrative
workflow involving document generation, file management, and
personalized email communication.
