# Installation and Setup Instructions

## Prerequisites

Before setting up the Expense Management System, ensure you have the following installed on your machine:

1. Microsoft Excel (Version 2013 or Later)
2. Git (Version 2.3.9) - For version control and managing the repository.

## Clone the Repository

To get started, clone this repository to your local machine using the following command in your terminal or Git Bash:

```
git clone https://github.com/IsraelVow/Expense-Management-System-VBA-Excel-Project-.git
```

## Opening the Excel-Based Project

1. Locate the cloned repository folder on your local machine.

2. In the "Expense Mgt sys" folder, you will find the Excel-based project file named "ExpenseMgtSystem.xlsm".

3. Double-click on "ExpenseMgtSystem.xlsm" to open the Excel project.

## Enable Macros

Upon opening the Excel project, you may encounter a security warning related to macros. To enable macros and ensure the proper functioning of the userform and backend functionalities, follow these steps:

1. Click on "Enable Macros" or "Enable Content" when prompted by Excel.

2. Go to the "Developer" tab on the Excel ribbon. If you don't see the "Developer" tab, you can enable it by going to Excel Options > Customize Ribbon > Check "Developer" in the right column > Click "OK."

3. In the "Code" group on the "Developer" tab, click on "Visual Basic" to open the Visual Basic for Applications (VBA) editor.

4. In the VBA editor, click on "Tools" > "References."

5. In the "References" window, make sure the following references are checked:

   - Microsoft ActiveX Data Objects 16.0 Library
   - Microsoft Forms 2.0 Object Library

6. Click "OK" to close the "References" window.

## Running the Expense Management System

1. In the Excel project, navigate to the "Interface" worksheet, where you will find buttons to interact with the Expense Management System.

2. Click on the "Expense Entry" button to access the userform for entering expenses.

3. Use the userform to enter expense details, and click "Save" to add the record to the backend.

4. To view the database and access analytics, click on the respective buttons in the "Interface" worksheet.

## Importing Expense Records

To import expense records from another spreadsheet, follow these steps:

1. Prepare a separate Excel file with expense records in a similar format to the template provided.

2. In the userform, click on the "Import Data" button.

3. Browse and select the Excel file containing the expense records you wish to import.

4. The system will import the data and update the backend accordingly.

## Additional Notes

- The dataset used in this project is synthesized and does not contain any financial or sensitive company information. All privacy and authorization considerations have been taken into account.

- For further assistance or inquiries, feel free to contact [Israel Josiah](israeljvow@gmail.com) or raise an issue on the [Israel Josiah](https://github.com/IsraelVow?tab=repositories).

---

By [Israel Josiah](https://www.linkedin.com/in/israeljosiah/) | LinkedIn: [Israel Josiah](https://www.linkedin.com/in/israeljosiah/)

In this "Installation and Setup Instructions" file, I have provided step-by-step guidance on setting up the project, enabling macros, running the system, and importing data. The file includes a "Prerequisites" section to ensure users have the required software installed and a note on the dataset's nature to address any privacy concerns. The contact information at the end allows users to reach out for further assistance or clarifications.