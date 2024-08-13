import os
import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QMenuBar, QMenu, QAction, QWidget, QVBoxLayout,
    QDialog, QLabel, QLineEdit, QPushButton, QFormLayout, QMessageBox, QTableWidget,
    QTableWidgetItem, QDateEdit, QComboBox, QDialogButtonBox, QGridLayout, QHBoxLayout
)
from PyQt5.QtCore import QDate, Qt
from fpdf import FPDF
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# Ensure data directory exists
os.makedirs("data", exist_ok=True)

class YouFish2GoRestaurantCoLLC(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("YouFish2Go Restaurant Co L.L.C")
        self.setGeometry(100, 100, 1000, 700)

        # Initialize data storage
        self.employees = pd.DataFrame(columns=['Name', 'Nationality', 'Designation', 'Basic Pay', 'Housing Allowance', 'Transportation Allowance'])
        self.sales = pd.DataFrame(columns=['Date', 'Amount', 'Type'])
        self.expenses = pd.DataFrame(columns=['Date', 'Amount', 'Category'])
        self.purchases = pd.DataFrame(columns=['Date', 'Company', 'Payment Type', 'Amount', 'Invoice Number', 'Remaining Balance'])
        self.accounts_payable = pd.DataFrame(columns=['Date', 'Company', 'Amount', 'Invoice Number', 'Remaining Balance'])
        self.accounts_receivable = pd.DataFrame(columns=['Date', 'Customer', 'Amount'])
        self.payslips = []
        self.advance_salaries = []

        # Load data from Excel files
        self.load_all_data()

        # Create Menu Bar
        self.create_menu_bar()

        # UI Components
        self.create_widgets()

    def create_menu_bar(self):
        menubar = self.menuBar()

        employee_menu = menubar.addMenu("Employee")
        add_employee_action = QAction("Add Employee", self)
        add_employee_action.triggered.connect(self.add_employee)
        employee_menu.addAction(add_employee_action)

        employee_list_action = QAction("Employee List", self)
        employee_list_action.triggered.connect(self.show_employee_list)
        employee_menu.addAction(employee_list_action)

        delete_employee_action = QAction("Delete Employee", self)
        delete_employee_action.triggered.connect(self.delete_employee)
        employee_menu.addAction(delete_employee_action)

        payslip_menu = menubar.addMenu("Payslip")
        generate_payslip_action = QAction("Generate Payslip", self)
        generate_payslip_action.triggered.connect(self.generate_salary_slip_page)
        payslip_menu.addAction(generate_payslip_action)

        generate_advance_action = QAction("Generate Advance Salary Slip", self)
        generate_advance_action.triggered.connect(self.generate_advance_salary_slip_page)
        payslip_menu.addAction(generate_advance_action)

        list_payslip_action = QAction("List Generated Payslips", self)
        list_payslip_action.triggered.connect(self.list_generated_payslips)
        payslip_menu.addAction(list_payslip_action)

        list_advance_action = QAction("List Generated Advance Payslips", self)
        list_advance_action.triggered.connect(self.list_generated_advance_salaries)
        payslip_menu.addAction(list_advance_action)

        transactions_menu = menubar.addMenu("Transactions")
        add_sales_action = QAction("Add Sales", self)
        add_sales_action.triggered.connect(self.add_sales)
        transactions_menu.addAction(add_sales_action)

        add_expense_action = QAction("Add Expense", self)
        add_expense_action.triggered.connect(self.add_expense)
        transactions_menu.addAction(add_expense_action)

        add_purchase_action = QAction("Add Purchase", self)
        add_purchase_action.triggered.connect(self.add_purchase)
        transactions_menu.addAction(add_purchase_action)

        reports_menu = menubar.addMenu("Reports")
        daily_sales_action = QAction("Daily Sales Report", self)
        daily_sales_action.triggered.connect(self.daily_sales_report)
        reports_menu.addAction(daily_sales_action)

        daily_expense_action = QAction("Daily Expense Report", self)
        daily_expense_action.triggered.connect(self.daily_expense_report)
        reports_menu.addAction(daily_expense_action)

        daily_purchase_action = QAction("Daily Purchase Report", self)
        daily_purchase_action.triggered.connect(self.daily_purchase_report)
        reports_menu.addAction(daily_purchase_action)

        custom_sales_action = QAction("Custom Sales Report", self)
        custom_sales_action.triggered.connect(self.generate_custom_sales_report)
        reports_menu.addAction(custom_sales_action)

        custom_expense_action = QAction("Custom Expense Report", self)
        custom_expense_action.triggered.connect(self.generate_custom_expense_report)
        reports_menu.addAction(custom_expense_action)

        custom_profit_loss_action = QAction("Custom Profit and Loss Report", self)
        custom_profit_loss_action.triggered.connect(self.generate_custom_profit_loss_report)
        reports_menu.addAction(custom_profit_loss_action)

        accounts_menu = menubar.addMenu("Accounts")
        accounts_payable_action = QAction("Accounts Payable", self)
        accounts_payable_action.triggered.connect(self.list_accounts_payable)
        accounts_menu.addAction(accounts_payable_action)

        accounts_receivable_action = QAction("Accounts Receivable", self)
        accounts_receivable_action.triggered.connect(self.list_accounts_receivable)
        accounts_menu.addAction(accounts_receivable_action)

        other_menu = menubar.addMenu("Other")
        profit_loss_action = QAction("Profit and Loss Statement", self)
        profit_loss_action.triggered.connect(self.profit_loss_statement)
        other_menu.addAction(profit_loss_action)

        dashboard_action = QAction("Dashboard", self)
        dashboard_action.triggered.connect(self.create_dashboard)
        other_menu.addAction(dashboard_action)

    def create_widgets(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(20)

        title_label = QLabel("YouFish2Go Restaurant Co L.L.C", self)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: blue;")
        main_layout.addWidget(title_label, alignment=Qt.AlignCenter)

        # Add Quick Action Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(20)

        add_employee_button = QPushButton("Add Employee", self)
        add_employee_button.clicked.connect(self.add_employee)
        button_layout.addWidget(add_employee_button)

        add_sales_button = QPushButton("Add Sales", self)
        add_sales_button.clicked.connect(self.add_sales)
        button_layout.addWidget(add_sales_button)

        add_expense_button = QPushButton("Add Expense", self)
        add_expense_button.clicked.connect(self.add_expense)
        button_layout.addWidget(add_expense_button)

        view_reports_button = QPushButton("View Reports", self)
        view_reports_button.clicked.connect(self.generate_custom_sales_report)
        button_layout.addWidget(view_reports_button)

        main_layout.addLayout(button_layout)

        # Add Summary Information
        summary_layout = QGridLayout()
        summary_layout.setSpacing(20)

        total_sales = self.sales['Amount'].sum()
        total_expenses = self.expenses['Amount'].sum()
        total_purchases = self.purchases['Amount'].sum()

        total_sales_label = QLabel(f"<h2>Total Sales: AED {total_sales}</h2>", self)
        total_expenses_label = QLabel(f"<h2>Total Expenses: AED {total_expenses}</h2>", self)
        total_purchases_label = QLabel(f"<h2>Total Purchases: AED {total_purchases}</h2>", self)

        summary_layout.addWidget(total_sales_label, 0, 0)
        summary_layout.addWidget(total_expenses_label, 0, 1)
        summary_layout.addWidget(total_purchases_label, 1, 0, 1, 2)

        main_layout.addLayout(summary_layout)
        summary_layout.setAlignment(Qt.AlignCenter)

        # Add a Chart
        chart_layout = QVBoxLayout()
        chart_layout.setContentsMargins(0, 0, 0, 0)
        chart_layout.setSpacing(10)

        # Ensure valid data for the pie chart
        if total_sales > 0 or total_expenses > 0:
            fig, ax = plt.subplots()
            ax.pie([total_sales, total_expenses], labels=['Sales', 'Expenses'], autopct='%1.1f%%')
            ax.axis('equal')
            canvas = FigureCanvas(fig)
            chart_layout.addWidget(canvas, alignment=Qt.AlignCenter)
        else:
            empty_chart_label = QLabel("No sales or expenses data available for the chart.")
            chart_layout.addWidget(empty_chart_label, alignment=Qt.AlignCenter)

        main_layout.addLayout(chart_layout)

        # Display Logo
        logo_path = os.path.join("data", "company_logo.png")
        if os.path.exists(logo_path):
            try:
                from PyQt5.QtGui import QPixmap
                logo = QLabel(self)
                pixmap = QPixmap(logo_path).scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo.setPixmap(pixmap)
                main_layout.addWidget(logo, alignment=Qt.AlignCenter)
            except Exception as e:
                print(f"Error loading logo: {e}")
            central_widget = QWidget()
            self.setCentralWidget(central_widget)
            main_layout = QVBoxLayout(central_widget)
            main_layout.setContentsMargins(10, 10, 10, 10)
            main_layout.setSpacing(20)

            title_label = QLabel("YouFish2Go Restaurant Co L.L.C", self)
            title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: blue;")
            main_layout.addWidget(title_label, alignment=Qt.AlignCenter)

            # Add Quick Action Buttons
            button_layout = QHBoxLayout()
            button_layout.setSpacing(20)

            add_employee_button = QPushButton("Add Employee", self)
            add_employee_button.clicked.connect(self.add_employee)
            button_layout.addWidget(add_employee_button)

            add_sales_button = QPushButton("Add Sales", self)
            add_sales_button.clicked.connect(self.add_sales)
            button_layout.addWidget(add_sales_button)

            add_expense_button = QPushButton("Add Expense", self)
            add_expense_button.clicked.connect(self.add_expense)
            button_layout.addWidget(add_expense_button)

            view_reports_button = QPushButton("View Reports", self)
            view_reports_button.clicked.connect(self.generate_custom_sales_report)
            button_layout.addWidget(view_reports_button)

            main_layout.addLayout(button_layout)

            # Add Summary Information
            summary_layout = QGridLayout()
            summary_layout.setSpacing(20)

            total_sales = self.sales['Amount'].sum()
            total_expenses = self.expenses['Amount'].sum()
            total_purchases = self.purchases['Amount'].sum()

            total_sales_label = QLabel(f"<h2>Total Sales: AED {total_sales}</h2>", self)
            total_expenses_label = QLabel(f"<h2>Total Expenses: AED {total_expenses}</h2>", self)
            total_purchases_label = QLabel(f"<h2>Total Purchases: AED {total_purchases}</h2>", self)

            summary_layout.addWidget(total_sales_label, 0, 0)
            summary_layout.addWidget(total_expenses_label, 0, 1)
            summary_layout.addWidget(total_purchases_label, 1, 0, 1, 2)

            main_layout.addLayout(summary_layout)
            summary_layout.setAlignment(Qt.AlignCenter)

            # Add a Chart
            chart_layout = QVBoxLayout()
            chart_layout.setContentsMargins(0, 0, 0, 0)
            chart_layout.setSpacing(10)

            if not (pd.isna(total_sales) or pd.isna(total_expenses)):
                fig, ax = plt.subplots()
                ax.pie([total_sales, total_expenses], labels=['Sales', 'Expenses'], autopct='%1.1f%%')
                ax.axis('equal')
                canvas = FigureCanvas(fig)
                chart_layout.addWidget(canvas, alignment=Qt.AlignCenter)

            main_layout.addLayout(chart_layout)

            # Display Logo
            logo_path = os.path.join("data", "company_logo.png")
            if os.path.exists(logo_path):
                try:
                    from PyQt5.QtGui import QPixmap
                    logo = QLabel(self)
                    pixmap = QPixmap(logo_path)
                    logo.setPixmap(pixmap)
                    main_layout.addWidget(logo, alignment=Qt.AlignCenter)
                except Exception as e:
                    print(f"Error loading logo: {e}")

    def load_from_excel(self, filename):
        filepath = os.path.join("data", filename)
        if os.path.exists(filepath):
            return pd.read_excel(filepath)
        return pd.DataFrame()

    def load_all_data(self):
        self.employees = self.load_from_excel('employees.xlsx')
        self.sales = self.load_from_excel('sales.xlsx')
        self.expenses = self.load_from_excel('expenses.xlsx')
        self.purchases = self.load_from_excel('purchases.xlsx')
        self.accounts_payable = self.load_from_excel('accounts_payable.xlsx')
        self.accounts_receivable = self.load_from_excel('accounts_receivable.xlsx')
        self.payslips = self.load_from_excel('payslips.xlsx').to_dict('records')
        self.advance_salaries = self.load_from_excel('advance_salaries.xlsx').to_dict('records')

        self.check_and_rename_columns(self.employees, 'employees', ['Name', 'Nationality', 'Designation', 'Basic Pay', 'Housing Allowance', 'Transportation Allowance'])
        self.check_and_rename_columns(self.sales, 'sales', ['Date', 'Amount', 'Type'])
        self.check_and_rename_columns(self.expenses, 'expenses', ['Date', 'Amount', 'Category'])
        self.check_and_rename_columns(self.purchases, 'purchases', ['Date', 'Company', 'Payment Type', 'Amount', 'Invoice Number', 'Remaining Balance'])
        self.check_and_rename_columns(self.accounts_payable, 'accounts_payable', ['Date', 'Company', 'Amount', 'Invoice Number', 'Remaining Balance'])
        self.check_and_rename_columns(self.accounts_receivable, 'accounts_receivable', ['Date', 'Customer', 'Amount'])

    def check_and_rename_columns(self, df, df_name, expected_columns):
        if set(expected_columns).issubset(df.columns):
            df = df[expected_columns]
        else:
            print(f"Warning: {df_name} DataFrame is missing expected columns. Available columns: {list(df.columns)}")
            df = pd.DataFrame(columns=expected_columns)
        setattr(self, df_name, df)

    def add_employee(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Employee")
        layout = QFormLayout(dialog)

        name_entry = QLineEdit(dialog)
        layout.addRow("Name", name_entry)

        nationality_entry = QLineEdit(dialog)
        layout.addRow("Nationality", nationality_entry)

        designation_entry = QLineEdit(dialog)
        layout.addRow("Designation", designation_entry)

        basic_pay_entry = QLineEdit(dialog)
        layout.addRow("Basic Pay", basic_pay_entry)

        housing_allowance_entry = QLineEdit(dialog)
        layout.addRow("Housing Allowance", housing_allowance_entry)

        transportation_allowance_entry = QLineEdit(dialog)
        layout.addRow("Transportation Allowance", transportation_allowance_entry)

        save_button = QPushButton("Save", dialog)
        save_button.clicked.connect(lambda: self.save_employee(dialog, name_entry, nationality_entry, designation_entry, basic_pay_entry, housing_allowance_entry, transportation_allowance_entry))
        layout.addWidget(save_button)
        dialog.exec_()

    def save_employee(self, dialog, name_entry, nationality_entry, designation_entry, basic_pay_entry, housing_allowance_entry, transportation_allowance_entry):
        new_employee = {
            'Name': name_entry.text(),
            'Nationality': nationality_entry.text(),
            'Designation': designation_entry.text(),
            'Basic Pay': float(basic_pay_entry.text() or 0),
            'Housing Allowance': float(housing_allowance_entry.text() or 0),
            'Transportation Allowance': float(transportation_allowance_entry.text() or 0)
        }
        self.employees = pd.concat([self.employees, pd.DataFrame([new_employee])], ignore_index=True)
        self.save_to_excel('employees.xlsx', self.employees)
        dialog.accept()
        self.show_employee_list()

    def show_employee_list(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Employee List")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(self.employees.shape[0], self.employees.shape[1], self)
        table.setHorizontalHeaderLabels(self.employees.columns)
        for i in range(self.employees.shape[0]):
            for j in range(self.employees.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(self.employees.iat[i, j])))

        layout.addWidget(table)
        dialog.exec_()

    def delete_employee(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Delete Employee")
        layout = QFormLayout(dialog)

        name_entry = QLineEdit(dialog)
        layout.addRow("Name", name_entry)

        delete_button = QPushButton("Delete", dialog)
        delete_button.clicked.connect(lambda: self.confirm_delete_employee(dialog, name_entry))
        layout.addWidget(delete_button)

        dialog.exec_()

    def confirm_delete_employee(self, dialog, name_entry):
        name = name_entry.text()
        self.employees = self.employees[self.employees['Name'] != name]
        self.save_to_excel('employees.xlsx', self.employees)
        dialog.accept()
        self.show_employee_list()

    def save_to_excel(self, filename, dataframe):
        filepath = os.path.join("data", filename)
        dataframe.to_excel(filepath, index=False)

    def generate_salary_slip_page(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Generate Salary Slip")
        layout = QFormLayout(dialog)

        employee_combobox = QComboBox(dialog)
        employee_combobox.addItems(self.employees['Name'].tolist())
        layout.addRow("Select Employee", employee_combobox)

        year_entry = QLineEdit(dialog)
        layout.addRow("Year", year_entry)

        month_entry = QLineEdit(dialog)
        layout.addRow("Month", month_entry)

        deductions_entry = QLineEdit(dialog)
        layout.addRow("Deductions (Optional)", deductions_entry)

        reason_entry = QLineEdit(dialog)
        layout.addRow("Reason for Deduction (Optional)", reason_entry)

        generate_button = QPushButton("Generate Salary Slip", dialog)
        generate_button.clicked.connect(lambda: self.generate_slip(dialog, employee_combobox, year_entry, month_entry, deductions_entry, reason_entry))
        layout.addWidget(generate_button)

        dialog.exec_()

    def generate_slip(self, dialog, employee_combobox, year_entry, month_entry, deductions_entry, reason_entry):
        selected_employee_name = employee_combobox.currentText()
        year = year_entry.text()
        month = month_entry.text()
        deductions = deductions_entry.text()
        reason = reason_entry.text()

        if not selected_employee_name or not year or not month:
            QMessageBox.warning(self, "Warning", "Please fill all required fields")
            return

        selected_employee = self.employees[self.employees['Name'] == selected_employee_name].iloc[0]
        name = selected_employee['Name']
        nationality = selected_employee['Nationality']
        designation = selected_employee['Designation']
        basic_pay = selected_employee['Basic Pay']
        housing_allowance = selected_employee['Housing Allowance']
        transportation_allowance = selected_employee['Transportation Allowance']

        if deductions:
            deductions = float(deductions)
        else:
            deductions = 0.0

        advance_salary_deducted = sum(
            adv['Advance Salary'] for adv in self.advance_salaries if adv['Employee'] == name and adv['Year'] == year and adv['Month'] == month
        )

        total_pay = basic_pay + housing_allowance + transportation_allowance - deductions - advance_salary_deducted
        creation_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        logo_path = os.path.join("data", "company_logo.png")
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=8, w=50)

        pdf.cell(200, 10, txt="Salary Slip", ln=True, align='C')
        pdf.cell(200, 10, txt="YouFish2Go Restaurant Co L.L.C", ln=True, align='C')
        pdf.cell(200, 10, txt="Al Rayees Shopping Center Shop No : 07", ln=True, align='C')
        pdf.cell(200, 10, txt="Landline: 042718736", ln=True, align='C')
        pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
        pdf.cell(200, 10, txt=f"Year: {year}", ln=True)
        pdf.cell(200, 10, txt=f"Month: {month}", ln=True)
        pdf.cell(200, 10, txt=f"Designation: {designation}", ln=True)
        pdf.cell(200, 10, txt=f"Basic Pay: AED {basic_pay}", ln=True)
        pdf.cell(200, 10, txt=f"Housing Allowance: AED {housing_allowance}", ln=True)
        pdf.cell(200, 10, txt=f"Transportation Allowance: AED {transportation_allowance}", ln=True)
        pdf.cell(200, 10, txt=f"Advance Salary Deducted: AED {advance_salary_deducted}", ln=True)
        pdf.cell(200, 10, txt=f"Deductions: AED {deductions}", ln=True)
        pdf.cell(200, 10, txt=f"Reason for Deduction: {reason}", ln=True)
        pdf.cell(200, 10, txt=f"Total Pay: AED {total_pay}", ln=True)
        pdf.cell(200, 10, ln=True)
        pdf.cell(200, 10, txt="Employee Signature: ___________________________", ln=True)

        pdf_filename = f"{name}_salary_slip_{year}_{month}.pdf"
        pdf.output(os.path.join("data", pdf_filename))

        new_payslip = {'Filename': pdf_filename, 'Employee': name, 'Creation Date': creation_datetime}
        self.payslips.append(new_payslip)
        self.save_to_excel('payslips.xlsx', pd.DataFrame(self.payslips))
        dialog.accept()
        QMessageBox.information(self, "Success", f"Salary slip for {name} generated successfully!")

    def generate_advance_salary_slip_page(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Generate Advance Salary Slip")
        layout = QFormLayout(dialog)

        employee_combobox = QComboBox(dialog)
        employee_combobox.addItems(self.employees['Name'].tolist())
        layout.addRow("Select Employee", employee_combobox)

        advance_amount_entry = QLineEdit(dialog)
        layout.addRow("Advance Amount", advance_amount_entry)

        generate_button = QPushButton("Generate Advance Salary Slip", dialog)
        generate_button.clicked.connect(lambda: self.generate_advance_slip(dialog, employee_combobox, advance_amount_entry))
        layout.addWidget(generate_button)

        dialog.exec_()

    def generate_advance_slip(self, dialog, employee_combobox, advance_amount_entry):
        selected_employee_name = employee_combobox.currentText()
        advance_amount = advance_amount_entry.text()

        if not selected_employee_name or not advance_amount:
            QMessageBox.warning(self, "Warning", "Please fill all required fields")
            return

        selected_employee = self.employees[self.employees['Name'] == selected_employee_name].iloc[0]
        name = selected_employee['Name']
        advance_amount = float(advance_amount)
        creation_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        logo_path = os.path.join("data", "company_logo.png")
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=8, w=50)

        pdf.cell(200, 10, txt="Advance Salary Slip", ln=True, align='C')
        pdf.cell(200, 10, txt="YouFish2Go Restaurant Co L.L.C", ln=True, align='C')
        pdf.cell(200, 10, txt="Al Rayees Shopping Center Shop No : 07", ln=True, align='C')
        pdf.cell(200, 10, txt="Landline: 042718736", ln=True, align='C')
        pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
        pdf.cell(200, 10, txt=f"Advance Amount: AED {advance_amount}", ln=True)

        pdf_filename = f"{name}_advance_salary_slip_{creation_datetime}.pdf"
        pdf.output(os.path.join("data", pdf_filename))
        new_advance_slip = {'Filename': pdf_filename, 'Employee': name, 'Creation Date': creation_datetime}
        self.advance_salaries.append(new_advance_slip)
        self.save_to_excel('advance_salaries.xlsx', pd.DataFrame(self.advance_salaries))
        dialog.accept()
        QMessageBox.information(self, "Success", f"Advance salary slip for {name} generated successfully!")

    def list_generated_advance_salaries(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("List of Generated Advance Salary Slips")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(len(self.advance_salaries), 3, self)
        table.setHorizontalHeaderLabels(['Filename', 'Employee', 'Creation Date'])
        for i, slip in enumerate(self.advance_salaries):
            table.setItem(i, 0, QTableWidgetItem(slip['Filename']))
            table.setItem(i, 1, QTableWidgetItem(slip['Employee']))
            table.setItem(i, 2, QTableWidgetItem(slip['Creation Date']))

        layout.addWidget(table)
        dialog.exec_()

    def list_generated_payslips(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("List of Generated Payslips")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(len(self.payslips), 3, self)
        table.setHorizontalHeaderLabels(['Filename', 'Employee', 'Creation Date'])
        for i, slip in enumerate(self.payslips):
            table.setItem(i, 0, QTableWidgetItem(slip['Filename']))
            table.setItem(i, 1, QTableWidgetItem(slip['Employee']))
            table.setItem(i, 2, QTableWidgetItem(slip['Creation Date']))

        layout.addWidget(table)
        dialog.exec_()

    def add_sales(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Sales")
        layout = QFormLayout(dialog)

        sales_date_entry = QDateEdit(calendarPopup=True)
        sales_date_entry.setDate(QDate.currentDate())
        layout.addRow("Sales Date", sales_date_entry)

        cash_sales_entry = QLineEdit(dialog)
        layout.addRow("Cash Sales Amount", cash_sales_entry)

        credit_sales_entry = QLineEdit(dialog)
        layout.addRow("Credit Card Sales Amount", credit_sales_entry)

        save_button = QPushButton("Save", dialog)
        save_button.clicked.connect(lambda: self.save_sales(dialog, sales_date_entry, cash_sales_entry, credit_sales_entry))
        layout.addWidget(save_button)

        dialog.exec_()

    def save_sales(self, dialog, sales_date_entry, cash_sales_entry, credit_sales_entry):
        sales_date = sales_date_entry.date().toString("yyyy-MM-dd")
        cash_sales_amount = float(cash_sales_entry.text() or 0)
        credit_sales_amount = float(credit_sales_entry.text() or 0)

        new_cash_sales = {'Date': sales_date, 'Amount': cash_sales_amount, 'Type': 'Cash'}
        new_credit_sales = {'Date': sales_date, 'Amount': credit_sales_amount, 'Type': 'Credit Card'}

        self.sales = pd.concat([self.sales, pd.DataFrame([new_cash_sales, new_credit_sales])], ignore_index=True)
        self.save_to_excel('sales.xlsx', self.sales)
        dialog.accept()
        QMessageBox.information(self, "Success", "Sales entry saved successfully!")

    def add_expense(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Expense")
        layout = QFormLayout(dialog)

        expense_date_entry = QDateEdit(calendarPopup=True)
        expense_date_entry.setDate(QDate.currentDate())
        layout.addRow("Expense Date", expense_date_entry)

        expense_amount_entry = QLineEdit(dialog)
        layout.addRow("Expense Amount", expense_amount_entry)

        expense_category_combobox = QComboBox(dialog)
        expense_category_combobox.addItems(["Rent", "Utilities", "Supplies", "Salaries", "Marketing"])
        layout.addRow("Expense Category", expense_category_combobox)

        save_button = QPushButton("Save", dialog)
        save_button.clicked.connect(lambda: self.save_expense(dialog, expense_date_entry, expense_amount_entry, expense_category_combobox))
        layout.addWidget(save_button)

        dialog.exec_()

    def save_expense(self, dialog, expense_date_entry, expense_amount_entry, expense_category_combobox):
        expense_date = expense_date_entry.date().toString("yyyy-MM-dd")
        expense_amount = float(expense_amount_entry.text() or 0)
        expense_category_value = expense_category_combobox.currentText()

        new_expense = {'Date': expense_date, 'Amount': expense_amount, 'Category': expense_category_value}
        self.expenses = pd.concat([self.expenses, pd.DataFrame([new_expense])], ignore_index=True)
        self.save_to_excel('expenses.xlsx', self.expenses)
        dialog.accept()
        QMessageBox.information(self, "Success", "Expense entry saved successfully!")

    def add_purchase(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Purchase")
        layout = QFormLayout(dialog)

        purchase_date_entry = QDateEdit(calendarPopup=True)
        purchase_date_entry.setDate(QDate.currentDate())
        layout.addRow("Purchase Date", purchase_date_entry)

        company_entry = QLineEdit(dialog)
        layout.addRow("Company Name", company_entry)

        payment_type_combobox = QComboBox(dialog)
        payment_type_combobox.addItems(["Cash", "Credit"])
        layout.addRow("Payment Type", payment_type_combobox)

        amount_entry = QLineEdit(dialog)
        layout.addRow("Amount", amount_entry)

        invoice_entry = QLineEdit(dialog)
        layout.addRow("Invoice Number", invoice_entry)

        save_button = QPushButton("Save", dialog)
        save_button.clicked.connect(lambda: self.save_purchase(dialog, purchase_date_entry, company_entry, payment_type_combobox, amount_entry, invoice_entry))
        layout.addWidget(save_button)

        dialog.exec_()

    def save_purchase(self, dialog, purchase_date_entry, company_entry, payment_type_combobox, amount_entry, invoice_entry):
        purchase_date = purchase_date_entry.date().toString("yyyy-MM-dd")
        company_name = company_entry.text()
        payment_type_value = payment_type_combobox.currentText()
        amount = float(amount_entry.text() or 0)
        invoice_number = invoice_entry.text()

        new_purchase = {'Date': purchase_date, 'Company': company_name, 'Payment Type': payment_type_value, 'Amount': amount, 'Invoice Number': invoice_number, 'Remaining Balance': amount}
        self.purchases = pd.concat([self.purchases, pd.DataFrame([new_purchase])], ignore_index=True)
        self.save_to_excel('purchases.xlsx', self.purchases)

        if payment_type_value == "Cash":
            new_expense = {'Date': purchase_date, 'Amount': amount, 'Category': f'Purchase from {company_name}'}
            self.expenses = pd.concat([self.expenses, pd.DataFrame([new_expense])], ignore_index=True)
            self.save_to_excel('expenses.xlsx', self.expenses)
            self.generate_payment_slip(company_name, amount, amount)
        else:
            new_account_payable = {'Date': purchase_date, 'Company': company_name, 'Amount': amount, 'Invoice Number': invoice_number, 'Remaining Balance': amount}
            self.accounts_payable = pd.concat([self.accounts_payable, pd.DataFrame([new_account_payable])], ignore_index=True)
            self.save_to_excel('accounts_payable.xlsx', self.accounts_payable)

        dialog.accept()
        QMessageBox.information(self, "Success", "Purchase entry saved successfully!")

    def daily_sales_report(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Daily Sales Report")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(self)
        layout.addWidget(table)

        today = datetime.now().strftime("%Y-%m-%d")
        daily_sales = self.sales[self.sales['Date'] == today]

        table.setRowCount(daily_sales.shape[0])
        table.setColumnCount(daily_sales.shape[1])
        table.setHorizontalHeaderLabels(daily_sales.columns)
        for i in range(daily_sales.shape[0]):
            for j in range(daily_sales.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(daily_sales.iat[i, j])))

        dialog.exec_()

    def daily_expense_report(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Daily Expense Report")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(self)
        layout.addWidget(table)

        today = datetime.now().strftime("%Y-%m-%d")
        daily_expenses = self.expenses[self.expenses['Date'] == today]

        table.setRowCount(daily_expenses.shape[0])
        table.setColumnCount(daily_expenses.shape[1])
        table.setHorizontalHeaderLabels(daily_expenses.columns)
        for i in range(daily_expenses.shape[0]):
            for j in range(daily_expenses.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(daily_expenses.iat[i, j])))

        dialog.exec_()

    def daily_purchase_report(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Daily Purchase Report")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(self)
        layout.addWidget(table)

        today = datetime.now().strftime("%Y-%m-%d")
        daily_purchases = self.purchases[self.purchases['Date'] == today]

        table.setRowCount(daily_purchases.shape[0])
        table.setColumnCount(daily_purchases.shape[1])
        table.setHorizontalHeaderLabels(daily_purchases.columns)
        for i in range(daily_purchases.shape[0]):
            for j in range(daily_purchases.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(daily_purchases.iat(i, j))))

        dialog.exec_()

    def list_accounts_payable(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Accounts Payable")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(self.accounts_payable.shape[0], self.accounts_payable.shape[1], self)
        table.setHorizontalHeaderLabels(self.accounts_payable.columns)
        for i in range(self.accounts_payable.shape[0]):
            for j in range(self.accounts_payable.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(self.accounts_payable.iat[i, j])))

        layout.addWidget(table)

        def mark_as_paid():
            selected_items = table.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Warning", "Please select an item to mark as paid")
                return

            row = selected_items[0].row()
            values = [table.item(row, col).text() for col in range(table.columnCount() - 1)]  # Exclude the Remaining Balance column

            print("Selected values:", values)  # Debugging print

            date_val, company_val, amount_val, invoice_val = values
            # Convert date_val to datetime format for comparison
            try:
                date_val = pd.to_datetime(date_val).date()
            except Exception as e:
                print(f"Error converting date: {e}")

            print(f"Converted Date: {date_val}")  # Debugging print
            print(f"DataFrame Dates: {pd.to_datetime(self.accounts_payable['Date']).dt.date}")  # Debugging print
            print(f"DataFrame Invoice Numbers: {self.accounts_payable['Invoice Number']}")  # Debugging print

            # Compare each field individually and print the results
            date_match = (pd.to_datetime(self.accounts_payable['Date']).dt.date == date_val)
            company_match = (self.accounts_payable['Company'] == company_val)
            amount_match = (self.accounts_payable['Amount'].astype(str) == amount_val)
            invoice_match = (self.accounts_payable['Invoice Number'].astype(str).str.strip() == invoice_val)

            print(f"Date Match: {date_match}")
            print(f"Company Match: {company_match}")
            print(f"Amount Match: {amount_match}")
            print(f"Invoice Match: {invoice_match}")

            matching_rows = self.accounts_payable[date_match & company_match & amount_match & invoice_match]

            print("Matching rows:", matching_rows)  # Debugging print

            if not matching_rows.empty:
                index = matching_rows.index[0]
                remaining_balance = self.accounts_payable.at[index, 'Remaining Balance']

                payment_dialog = QDialog(self)
                payment_dialog.setWindowTitle("Mark as Paid")
                payment_layout = QFormLayout(payment_dialog)

                payment_layout.addRow(QLabel(f"Company: {values[1]}"))
                payment_layout.addRow(QLabel(f"Total Amount: AED {values[2]}"))
                payment_layout.addRow(QLabel(f"Remaining Balance: AED {remaining_balance}"))

                payment_amount_entry = QLineEdit(payment_dialog)
                payment_layout.addRow("Payment Amount", payment_amount_entry)

                def confirm_payment():
                    payment_amount = float(payment_amount_entry.text() or 0)
                    if payment_amount > remaining_balance:
                        QMessageBox.critical(self, "Error", "Payment amount exceeds remaining balance.")
                        return

                    new_balance = remaining_balance - payment_amount
                    if new_balance == 0:
                        self.accounts_payable = self.accounts_payable.drop(index)
                    else:
                        self.accounts_payable.at[index, 'Remaining Balance'] = new_balance

                    table.removeRow(row)
                    self.add_expense_from_payment(values[1], payment_amount)
                    self.save_to_excel('accounts_payable.xlsx', self.accounts_payable)
                    self.generate_payment_slip(values[1], values[2], new_balance)
                    QMessageBox.information(self, "Success", "Marked as paid and expense recorded.")
                    payment_dialog.accept()

                confirm_button = QPushButton("Confirm Payment", payment_dialog)
                confirm_button.clicked.connect(confirm_payment)
                payment_layout.addWidget(confirm_button)

                payment_dialog.exec_()

            else:
                QMessageBox.critical(self, "Error", "No matching entry found to mark as paid.")

        mark_as_paid_button = QPushButton("Mark as Paid", self)
        mark_as_paid_button.clicked.connect(mark_as_paid)
        layout.addWidget(mark_as_paid_button)

        dialog.exec_()

    def list_accounts_receivable(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Accounts Receivable")
        layout = QVBoxLayout(dialog)

        table = QTableWidget(self.accounts_receivable.shape[0], self.accounts_receivable.shape[1], self)
        table.setHorizontalHeaderLabels(self.accounts_receivable.columns)
        for i in range(self.accounts_receivable.shape[0]):
            for j in range(self.accounts_receivable.shape[1]):
                table.setItem(i, j, QTableWidgetItem(str(self.accounts_receivable.iat(i, j))))

        layout.addWidget(table)

        def mark_as_paid():
            selected_items = table.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Warning", "Please select an item to mark as paid")
                return

            row = selected_items[0].row()
            values = [table.item(row, col).text() for col in range(table.columnCount())]

            matching_rows = self.accounts_receivable[
                (self.accounts_receivable['Date'] == values[0]) &
                (self.accounts_receivable['Customer'] == values[1]) &
                (self.accounts_receivable['Amount'] == float(values[2]))
            ]
            if not matching_rows.empty:
                index = matching_rows.index[0]
                self.accounts_receivable = self.accounts_receivable.drop(index)
                table.removeRow(row)
                self.save_to_excel('accounts_receivable.xlsx', self.accounts_receivable)
                QMessageBox.information(self, "Success", "Marked as paid.")
            else:
                QMessageBox.critical(self, "Error", "No matching entry found to mark as paid.")

        mark_as_paid_button = QPushButton("Mark as Paid", self)
        mark_as_paid_button.clicked.connect(mark_as_paid)
        layout.addWidget(mark_as_paid_button)

        dialog.exec_()

    def add_expense_from_payment(self, company, amount):
        today = datetime.now().strftime("%Y-%m-%d")
        new_expense = {'Date': today, 'Amount': amount, 'Category': f'Payment to {company}'}
        self.expenses = pd.concat([self.expenses, pd.DataFrame([new_expense])], ignore_index=True)
        self.save_to_excel('expenses.xlsx', self.expenses)

    def generate_payment_slip(self, company, total_amount, remaining_balance):
        creation_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        logo_path = os.path.join("data", "company_logo.png")
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=8, w=50)

        pdf.cell(200, 10, txt="Payment Slip", ln=True, align='C')
        pdf.cell(200, 10, txt="YouFish2Go Restaurant Co L.L.C", ln=True, align='C')
        pdf.cell(200, 10, txt="Al Rayees Shopping Center Shop No : 07", ln=True, align='C')
        pdf.cell(200, 10, txt="Landline: 042718736", ln=True, align='C')
        pdf.cell(200, 10, txt=f"Company: {company}", ln=True)
        pdf.cell(200, 10, txt=f"Total Amount: AED {total_amount}", ln=True)
        pdf.cell(200, 10, txt=f"Remaining Balance: AED {remaining_balance}", ln=True)

        pdf_filename = f"{company}_payment_slip_{creation_datetime}.pdf"
        pdf.output(os.path.join("data", pdf_filename))

    def generate_custom_sales_report(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Custom Sales Report")
        layout = QFormLayout(dialog)

        start_date_entry = QDateEdit(calendarPopup=True)
        start_date_entry.setDate(QDate.currentDate())
        layout.addRow("Start Date", start_date_entry)

        end_date_entry = QDateEdit(calendarPopup=True)
        end_date_entry.setDate(QDate.currentDate())
        layout.addRow("End Date", end_date_entry)

        generate_button = QPushButton("Generate Report", dialog)
        generate_button.clicked.connect(lambda: self.generate_sales_report(dialog, start_date_entry, end_date_entry))
        layout.addWidget(generate_button)

        dialog.exec_()

    def generate_sales_report(self, dialog, start_date_entry, end_date_entry):
        start_date = start_date_entry.date().toString("yyyy-MM-dd")
        end_date = end_date_entry.date().toString("yyyy-MM-dd")

        report_data = self.sales[(self.sales['Date'] >= start_date) & (self.sales['Date'] <= end_date)]

        fig, ax = plt.subplots()
        report_data.groupby('Type').sum()['Amount'].plot(kind='bar', ax=ax)
        ax.set_title('Sales Report')
        ax.set_xlabel('Type')
        ax.set_ylabel('Amount')
        ax.grid(True)

        canvas = FigureCanvas(fig)
        canvas.draw()

        report_dialog = QDialog(self)
        report_dialog.setWindowTitle("Sales Report")
        report_layout = QVBoxLayout(report_dialog)
        report_layout.addWidget(canvas)
        report_dialog.exec_()

    def generate_custom_expense_report(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Custom Expense Report")
        layout = QFormLayout(dialog)

        start_date_entry = QDateEdit(calendarPopup=True)
        start_date_entry.setDate(QDate.currentDate())
        layout.addRow("Start Date", start_date_entry)

        end_date_entry = QDateEdit(calendarPopup=True)
        end_date_entry.setDate(QDate.currentDate())
        layout.addRow("End Date", end_date_entry)

        generate_button = QPushButton("Generate Report", dialog)
        generate_button.clicked.connect(lambda: self.generate_expense_report(dialog, start_date_entry, end_date_entry))
        layout.addWidget(generate_button)

        dialog.exec_()

    def generate_expense_report(self, dialog, start_date_entry, end_date_entry):
        start_date = start_date_entry.date().toString("yyyy-MM-dd")
        end_date = end_date_entry.date().toString("yyyy-MM-dd")

        report_data = self.expenses[(self.expenses['Date'] >= start_date) & (self.expenses['Date'] <= end_date)]

        fig, ax = plt.subplots()
        report_data.groupby('Category').sum()['Amount'].plot(kind='bar', ax=ax)
        ax.set_title('Expense Report')
        ax.set_xlabel('Category')
        ax.set_ylabel('Amount')
        ax.grid(True)

        canvas = FigureCanvas(fig)
        canvas.draw()

        report_dialog = QDialog(self)
        report_dialog.setWindowTitle("Expense Report")
        report_layout = QVBoxLayout(report_dialog)
        report_layout.addWidget(canvas)
        report_dialog.exec_()

    def generate_custom_profit_loss_report(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Custom Profit and Loss Report")
        layout = QFormLayout(dialog)

        start_date_entry = QDateEdit(calendarPopup=True)
        start_date_entry.setDate(QDate.currentDate())
        layout.addRow("Start Date", start_date_entry)

        end_date_entry = QDateEdit(calendarPopup=True)
        end_date_entry.setDate(QDate.currentDate())
        layout.addRow("End Date", end_date_entry)

        generate_button = QPushButton("Generate Report", dialog)
        generate_button.clicked.connect(lambda: self.generate_profit_loss_report(dialog, start_date_entry, end_date_entry))
        layout.addWidget(generate_button)

        dialog.exec_()

    def generate_profit_loss_report(self, dialog, start_date_entry, end_date_entry):
        start_date = start_date_entry.date().toString("yyyy-MM-dd")
        end_date = end_date_entry.date().toString("yyyy-MM-dd")

        sales_data = self.sales[(self.sales['Date'] >= start_date) & (self.sales['Date'] <= end_date)]
        expense_data = self.expenses[(self.expenses['Date'] >= start_date) & (self.expenses['Date'] <= end_date)]

        total_sales = sales_data['Amount'].sum()
        total_expenses = expense_data['Amount'].sum()
        profit_loss = total_sales - total_expenses

        fig, ax = plt.subplots()
        ax.bar(['Total Sales', 'Total Expenses', 'Profit/Loss'], [total_sales, total_expenses, profit_loss])
        ax.set_title('Profit and Loss Report')
        ax.set_xlabel('Category')
        ax.set_ylabel('Amount')
        ax.grid(True)

        canvas = FigureCanvas(fig)
        canvas.draw()

        report_dialog = QDialog(self)
        report_dialog.setWindowTitle("Profit and Loss Report")
        report_layout = QVBoxLayout(report_dialog)
        report_layout.addWidget(canvas)
        report_dialog.exec_()

    def profit_loss_statement(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Profit and Loss Statement")
        layout = QVBoxLayout(dialog)

        total_sales = self.sales['Amount'].sum()
        total_expenses = self.expenses['Amount'].sum()
        profit_loss = total_sales - total_expenses

        layout.addWidget(QLabel(f"Total Sales: AED {total_sales}"))
        layout.addWidget(QLabel(f"Total Expenses: AED {total_expenses}"))
        layout.addWidget(QLabel(f"Profit/Loss: AED {profit_loss}"))

        dialog.exec_()

    def create_dashboard(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Dashboard")
        layout = QVBoxLayout(dialog)

        total_sales = self.sales['Amount'].sum()
        total_expenses = self.expenses['Amount'].sum()
        total_purchases = self.purchases['Amount'].sum()
        total_accounts_payable = self.accounts_payable['Amount'].sum()
        total_accounts_receivable = self.accounts_receivable['Amount'].sum()

        layout.addWidget(QLabel(f"Total Sales: AED {total_sales}"))
        layout.addWidget(QLabel(f"Total Expenses: AED {total_expenses}"))
        layout.addWidget(QLabel(f"Total Purchases: AED {total_purchases}"))
        layout.addWidget(QLabel(f"Total Accounts Payable: AED {total_accounts_payable}"))
        layout.addWidget(QLabel(f"Total Accounts Receivable: AED {total_accounts_receivable}"))

        dialog.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = YouFish2GoRestaurantCoLLC()
    window.show()
    sys.exit(app.exec_())
