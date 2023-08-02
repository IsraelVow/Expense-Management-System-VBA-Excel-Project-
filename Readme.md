**Expense Management System - VBA Excel Project**

<div style="display: flex; flex-direction: row;">
  <img src="https://i.imgur.com/fyv3tVj.png" alt="Dashboard" width="400" style="margin-right: 20px;">
  <img src="https://i.imgur.com/yFxloFI.png" alt="Userform" width="400">
</div>


## Overview

The Expense Management System is a VBA Excel-based project designed to provide organizations with an efficient and user-friendly solution for managing expenses. The system allows users to enter, track, and analyze expenses through a custom userform and real-time analytics dashboard. It also offers the capability to import expense records from other spreadsheets and provides extensive data visualization for better decision-making.

## Features

1. **Userform Entry:** The system features a user-friendly userform that enables users to input expense details, such as date, category, expense name, amount, location, comments, and more.

2. **Real-Time Dashboard:** The analytics dashboard presents key insights into expense data using various charts and visualizations, including donut charts, line charts, waterfall charts, and battery charts, providing a comprehensive overview of expenses.

3. **Pivot Tables and Power Query:** The project utilizes Power Pivot and Power Query to create data models, relationships, and perform EDA (Exploratory Data Analysis) for in-depth expense analysis.

4. **Dependent Dropdowns:** The userform features dependent dropdowns, enabling users to select expense names based on the chosen category, ensuring accurate and relevant expense entries.

5. **Data Import and Export:** The system supports importing expenses from other spreadsheets, and users can export data to different formats, such as CSV, for seamless sharing and analysis.

6. **Security and Admin Privileges:** A System Admin section provides options to change passwords, set user permissions, and manage system settings, ensuring data security and user management.

7. **Cloud-Based Database (Future Enhancement):** The project has a vision for future improvements, including transitioning the database to a cloud-based solution, allowing real-time collaboration and simultaneous entry of expenses by multiple users.

## Technical Details

### Tools and Technologies Used

- VBA (Visual Basic for Applications)
- Microsoft Excel (Office 2019)

### Data Manipulation and Analysis

- Power Pivot for creating data models and relationships
- Power Query for ETL (Extract, Transform, Load) operations
- Pivot Tables for generating insights from expense data

### User Interface

- Userform for data entry with dependent dropdowns for category and expense name selection
- Analytics Dashboard with various charts and visualizations for expense insights

### Data Structure

- "fTransaction" Table for storing expense data with columns like Date, Category, Expense Name, Amount, Location, Comments, etc.
- "Dexpense_category" Table as a dimension table for expense categories
- "Dexpense_name" Table as a dimension table for individual expense names

## Continuous Improvement

The project aims to continuously improve and enhance its functionalities, including the following future enhancements:

1. Converting the Software to an EXE File: Creating an executable version of the Expense Management System for easy distribution and usage.
2. Exporting to Different File Formats: Adding the capability to export expense data to CSV and other formats for improved data sharing.
3. Cloud-Based Database (Future Enhancement): Moving the database to a cloud-based solution to enable real-time collaboration and simultaneous data entry.

## Usage

To use the Expense Management System, open the Excel workbook and navigate to the "Interface" sheet to access the userform for entering expenses. Use the "Analysis" sheet to view analytics and the "System Admin" section for managing user privileges.

## License

The Expense Management System is open-source and licensed under [MIT License](link-to-license).

## Contributors

- Israel Josiah https://github.com/IsraelVow - Lead Developer & Data Analyst

## Feedback and Contributions

We welcome your feedback and contributions to further improve the Expense Management System. Feel free to submit issues or pull requests on our [GitHub repository](https://github.com/IsraelVow?tab=repositories).

## Acknowledgments

We extend our gratitude to the open-source community for their valuable contributions and inspirations that helped shape this project.

---

By [Israel Josiah](https://github.com/IsraelVow) | LinkedIn: https://www.linkedin.com/in/israeljosiah/

For any inquiries or support, please contact us at (mailto:Israeljvow@gmail.com).
