"""
Advanced Task Management Excel System
- Multi-sheet workbook with automatic task routing
- Live dashboard with real-time task tracking
- Automatic categorization and assignment tracking
- Status monitoring and completion tracking
"""

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, timedelta
import os

class TaskManagementSystem:
    def __init__(self, filename='task_management.xlsx'):
        self.filename = filename
        self.wb = None
        self.sheets = {}
        
        # Task categories - customize as needed
        self.categories = [
            'Development',
            'Design',
            'Marketing',
            'Sales',
            'Support',
            'Admin',
            'Planning',
            'Research'
        ]
        
        # Priority levels
        self.priorities = ['Low', 'Medium', 'High', 'Critical']
        
        # Status options
        self.statuses = ['Not Started', 'In Progress', 'On Hold', 'Completed', 'Cancelled']
        
        self.setup_workbook()
    
    def setup_workbook(self):
        """Initialize or load workbook"""
        if os.path.exists(self.filename):
            self.wb = load_workbook(self.filename)
            print(f"📂 Loaded existing workbook: {self.filename}")
        else:
            self.wb = Workbook()
            self.wb.remove(self.wb.active)
            print(f"📄 Created new workbook: {self.filename}")
        
        # Create all necessary sheets
        self.create_dashboard()
        self.create_task_entry_form()
        self.create_category_sheets()
        self.create_team_view()
        self.create_analytics_sheet()
        self.create_archive_sheet()
        
        self.save()
    
    def create_dashboard(self):
        """Create live dashboard with KPIs and task overview"""
        if 'Dashboard' in self.wb.sheetnames:
            ws = self.wb['Dashboard']
        else:
            ws = self.wb.create_sheet('Dashboard', 0)
        
        self.sheets['Dashboard'] = ws
        
        # Clear existing content
        ws.delete_rows(1, ws.max_row)
        
        # Set column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 20
        ws.column_dimensions['G'].width = 15
        
        # Title
        ws['A1'] = '📊 LIVE TASK DASHBOARD'
        ws['A1'].font = Font(bold=True, size=20, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:G1')
        ws.row_dimensions[1].height = 35
        
        # Last Updated
        ws['A2'] = 'Last Updated:'
        ws['B2'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ws['B2'].font = Font(italic=True, color='7F8C8D')
        
        # KPI Section
        ws['A4'] = '🎯 KEY PERFORMANCE INDICATORS'
        ws['A4'].font = Font(bold=True, size=14, color='2C3E50')
        ws.merge_cells('A4:G4')
        
        # KPI Cards (will be populated with formulas)
        kpis = [
            ('A6', 'B6', 'Total Tasks', "=COUNTA('Task Entry'!A:A)-1"),
            ('C6', 'D6', 'Active Tasks', "=COUNTIFS('Task Entry'!F:F,\"In Progress\",'Task Entry'!F:F,\"Not Started\")"),
            ('E6', 'F6', 'Completed', "=COUNTIF('Task Entry'!F:F,\"Completed\")"),
            ('G6', 'G6', 'Overdue', "=COUNTIF('Task Entry'!E:E,\"<\"&TODAY())")
        ]
        
        for start_cell, end_cell, label, formula in kpis:
            # Label
            ws[start_cell] = label
            ws[start_cell].font = Font(bold=True, size=11, color='FFFFFF')
            ws[start_cell].fill = PatternFill(start_color='3498DB', end_color='3498DB', fill_type='solid')
            ws[start_cell].alignment = Alignment(horizontal='center', vertical='center')
            if start_cell != end_cell:
                ws.merge_cells(f'{start_cell}:{end_cell}')
            
            # Value (row below)
            value_cell = start_cell[0] + '7'
            ws[value_cell] = formula
            ws[value_cell].font = Font(bold=True, size=18, color='2C3E50')
            ws[value_cell].alignment = Alignment(horizontal='center', vertical='center')
            if start_cell != end_cell:
                end_value = end_cell[0] + '7'
                ws.merge_cells(f'{value_cell}:{end_value}')
        
        ws.row_dimensions[6].height = 25
        ws.row_dimensions[7].height = 35
        
        # Outstanding Tasks Section
        ws['A9'] = '⚠️ OUTSTANDING INCOMPLETE TASKS'
        ws['A9'].font = Font(bold=True, size=14, color='E74C3C')
        ws.merge_cells('A9:G9')
        
        # Outstanding Tasks Header
        headers = ['Task ID', 'Task Name', 'Category', 'Assigned To', 'Priority', 'Due Date', 'Status']
        for idx, header in enumerate(headers, 1):
            cell = ws.cell(row=10, column=idx)
            cell.value = header
            cell.font = Font(bold=True, color='FFFFFF', size=11)
            cell.fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        ws.row_dimensions[10].height = 25
        
        # Add formulas to pull incomplete tasks (rows 11-30)
        # This will show tasks that are NOT completed
        ws['A11'] = '=IF(ROWS(A$11:A11)>COUNTIFS(\'Task Entry\'!F:F,"<>Completed",\'Task Entry\'!F:F,"<>Cancelled"),"",INDEX(\'Task Entry\'!A:A,SMALL(IF(\'Task Entry\'!F$2:F$1000<>"Completed",IF(\'Task Entry\'!F$2:F$1000<>"Cancelled",ROW(\'Task Entry\'!F$2:F$1000))),ROWS(A$11:A11))))'
        
        # Instructions for the user
        ws['A32'] = '📝 HOW TO USE THIS DASHBOARD:'
        ws['A32'].font = Font(bold=True, size=12, color='2C3E50')
        
        instructions = [
            "1. This dashboard updates automatically when you add tasks",
            "2. All incomplete tasks are shown in the red table above",
            "3. KPIs update in real-time based on task status",
            "4. Go to 'Task Entry' sheet to add new tasks",
            "5. Tasks automatically route to category sheets",
            "6. Refresh (F9) to update live data"
        ]
        
        for idx, instruction in enumerate(instructions, 33):
            ws[f'A{idx}'] = instruction
            ws[f'A{idx}'].font = Font(size=10, color='7F8C8D')
        
        # Freeze panes
        ws.freeze_panes = 'A11'
        
        return ws
    
    def create_task_entry_form(self):
        """Create main task entry sheet"""
        if 'Task Entry' in self.wb.sheetnames:
            ws = self.wb['Task Entry']
        else:
            ws = self.wb.create_sheet('Task Entry', 1)
        
        self.sheets['Task Entry'] = ws
        
        # Clear existing headers if needed
        if ws.max_row == 1:
            # Set column widths
            widths = [10, 30, 15, 25, 12, 15, 15, 12, 40, 20]
            for idx, width in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(idx)].width = width
            
            # Headers
            headers = [
                'Task ID',
                'Task Name',
                'Category',
                'Assigned To',
                'Priority',
                'Due Date',
                'Status',
                'Progress %',
                'Description',
                'Created Date'
            ]
            
            for idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=idx)
                cell.value = header
                cell.font = Font(bold=True, color='FFFFFF', size=12)
                cell.fill = PatternFill(start_color='27AE60', end_color='27AE60', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            
            ws.row_dimensions[1].height = 30
            
            # Add data validation for Category
            category_validation = DataValidation(
                type="list",
                formula1=f'"{",".join(self.categories)}"',
                allow_blank=False
            )
            ws.add_data_validation(category_validation)
            category_validation.add('C2:C1000')
            
            # Add data validation for Priority
            priority_validation = DataValidation(
                type="list",
                formula1=f'"{",".join(self.priorities)}"',
                allow_blank=False
            )
            ws.add_data_validation(priority_validation)
            priority_validation.add('E2:E1000')
            
            # Add data validation for Status
            status_validation = DataValidation(
                type="list",
                formula1=f'"{",".join(self.statuses)}"',
                allow_blank=False
            )
            ws.add_data_validation(status_validation)
            status_validation.add('G2:G1000')
            
            # Add sample task
            sample_task = [
                '=ROW()-1',  # Auto ID
                'Sample Task - Delete Me',
                'Development',
                'John Doe',
                'Medium',
                (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d'),
                'Not Started',
                0,
                'This is a sample task. Delete this row and add your own tasks.',
                datetime.now().strftime('%Y-%m-%d')
            ]
            
            for idx, value in enumerate(sample_task, 1):
                ws.cell(row=2, column=idx, value=value)
        
        # Add conditional formatting for Priority
        priority_rules = {
            'Critical': 'C0392B',
            'High': 'E74C3C',
            'Medium': 'F39C12',
            'Low': '27AE60'
        }
        
        for priority, color in priority_rules.items():
            ws.conditional_formatting.add(
                'E2:E1000',
                CellIsRule(
                    operator='equal',
                    formula=[f'"{priority}"'],
                    fill=PatternFill(start_color=color, end_color=color, fill_type='solid'),
                    font=Font(bold=True, color='FFFFFF')
                )
            )
        
        # Add conditional formatting for Status
        status_colors = {
            'Completed': '27AE60',
            'In Progress': '3498DB',
            'On Hold': 'F39C12',
            'Not Started': '95A5A6',
            'Cancelled': '7F8C8D'
        }
        
        for status, color in status_colors.items():
            ws.conditional_formatting.add(
                'G2:G1000',
                CellIsRule(
                    operator='equal',
                    formula=[f'"{status}"'],
                    fill=PatternFill(start_color=color, end_color=color, fill_type='solid'),
                    font=Font(bold=True, color='FFFFFF')
                )
            )
        
        # Freeze header row
        ws.freeze_panes = 'A2'
        
        return ws
    
    def create_category_sheets(self):
        """Create detailed sheets for each category"""
        for category in self.categories:
            if category in self.wb.sheetnames:
                ws = self.wb[category]
            else:
                ws = self.wb.create_sheet(category)
            
            self.sheets[category] = ws
            
            # Set up headers
            if ws.max_row == 1:
                # Column widths
                widths = [10, 30, 25, 12, 15, 15, 12, 40, 20, 20, 40]
                for idx, width in enumerate(widths, 1):
                    ws.column_dimensions[get_column_letter(idx)].width = width
                
                headers = [
                    'Task ID',
                    'Task Name',
                    'Assigned To',
                    'Priority',
                    'Due Date',
                    'Status',
                    'Progress %',
                    'Description',
                    'Created Date',
                    'Completed Date',
                    'Notes'
                ]
                
                for idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=idx)
                    cell.value = header
                    cell.font = Font(bold=True, color='FFFFFF', size=11)
                    cell.fill = PatternFill(start_color='8E44AD', end_color='8E44AD', fill_type='solid')
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Add instruction
                ws['A2'] = f'Tasks from "{category}" category will automatically appear here'
                ws['A2'].font = Font(italic=True, color='7F8C8D')
                ws.merge_cells('A2:K2')
                
                ws.freeze_panes = 'A2'
        
        return self.sheets
    
    def create_team_view(self):
        """Create team member view showing their assigned tasks"""
        if 'Team View' in self.wb.sheetnames:
            ws = self.wb['Team View']
        else:
            ws = self.wb.create_sheet('Team View')
        
        self.sheets['Team View'] = ws
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 12
        
        # Title
        ws['A1'] = '👥 TEAM TASK VIEW'
        ws['A1'].font = Font(bold=True, size=16, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color='16A085', end_color='16A085', fill_type='solid')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:F1')
        
        # Headers
        headers = ['Assigned To', 'Task Name', 'Category', 'Priority', 'Due Date', 'Status']
        for idx, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=idx)
            cell.value = header
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='1ABC9C', end_color='1ABC9C', fill_type='solid')
            cell.alignment = Alignment(horizontal='center')
        
        ws.freeze_panes = 'A3'
        
        return ws
    
    def create_analytics_sheet(self):
        """Create analytics and reporting sheet"""
        if 'Analytics' in self.wb.sheetnames:
            ws = self.wb['Analytics']
        else:
            ws = self.wb.create_sheet('Analytics')
        
        self.sheets['Analytics'] = ws
        
        ws['A1'] = '📈 TASK ANALYTICS & REPORTS'
        ws['A1'].font = Font(bold=True, size=16, color='2C3E50')
        
        # Summary statistics
        ws['A3'] = 'Summary Statistics'
        ws['A3'].font = Font(bold=True, size=12)
        
        stats = [
            ('A5', 'Total Tasks:', "=COUNTA('Task Entry'!A:A)-1"),
            ('A6', 'Completed Tasks:', "=COUNTIF('Task Entry'!G:G,\"Completed\")"),
            ('A7', 'In Progress:', "=COUNTIF('Task Entry'!G:G,\"In Progress\")"),
            ('A8', 'Not Started:', "=COUNTIF('Task Entry'!G:G,\"Not Started\")"),
            ('A9', 'On Hold:', "=COUNTIF('Task Entry'!G:G,\"On Hold\")"),
            ('A10', 'Completion Rate:', "=IF(B5>0,B6/B5,0)"),
        ]
        
        for cell, label, formula in stats:
            ws[cell] = label
            ws[cell].font = Font(bold=True)
            value_cell = 'B' + cell[1:]
            ws[value_cell] = formula
            if 'Rate' in label:
                ws[value_cell].number_format = '0.00%'
        
        # Category breakdown
        ws['A12'] = 'Tasks by Category'
        ws['A12'].font = Font(bold=True, size=12)
        
        for idx, category in enumerate(self.categories, 14):
            ws[f'A{idx}'] = category
            ws[f'B{idx}'] = f'=COUNTIF(\'Task Entry\'!C:C,"{category}")'
        
        return ws
    
    def create_archive_sheet(self):
        """Create archive for completed tasks"""
        if 'Archive' in self.wb.sheetnames:
            ws = self.wb['Archive']
        else:
            ws = self.wb.create_sheet('Archive')
        
        self.sheets['Archive'] = ws
        
        ws['A1'] = '📦 COMPLETED TASKS ARCHIVE'
        ws['A1'].font = Font(bold=True, size=14, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color='7F8C8D', end_color='7F8C8D', fill_type='solid')
        ws.merge_cells('A1:J1')
        
        headers = ['Task ID', 'Task Name', 'Category', 'Assigned To', 'Priority', 
                   'Due Date', 'Completed Date', 'Duration (days)', 'Description', 'Notes']
        
        for idx, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=idx)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='BDC3C7', end_color='BDC3C7', fill_type='solid')
        
        return ws
    
    def add_task(self, task_data):
        """Add a new task to the system"""
        ws = self.sheets['Task Entry']
        
        # Find next empty row
        next_row = ws.max_row + 1
        
        # Add task data
        task_id = next_row - 1
        row_data = [
            task_id,
            task_data.get('name', ''),
            task_data.get('category', 'Admin'),
            task_data.get('assigned_to', ''),
            task_data.get('priority', 'Medium'),
            task_data.get('due_date', ''),
            task_data.get('status', 'Not Started'),
            task_data.get('progress', 0),
            task_data.get('description', ''),
            datetime.now().strftime('%Y-%m-%d')
        ]
        
        for idx, value in enumerate(row_data, 1):
            ws.cell(row=next_row, column=idx, value=value)
        
        print(f"✅ Added task: {task_data.get('name')} (ID: {task_id})")
        
        self.save()
        return task_id
    
    def save(self):
        """Save the workbook"""
        self.wb.save(self.filename)
        return self.filename

# Create and test the system
if __name__ == '__main__':
    print("\n🚀 Creating Task Management System...")
    print("=" * 60)
    
    # Create system
    tms = TaskManagementSystem('task_management.xlsx')
    
    # Add sample tasks
    sample_tasks = [
        {
            'name': 'Build authentication system',
            'category': 'Development',
            'assigned_to': 'Alice Smith',
            'priority': 'High',
            'due_date': (datetime.now() + timedelta(days=5)).strftime('%Y-%m-%d'),
            'status': 'In Progress',
            'progress': 30,
            'description': 'Implement JWT authentication for API'
        },
        {
            'name': 'Design landing page mockup',
            'category': 'Design',
            'assigned_to': 'Bob Johnson',
            'priority': 'Medium',
            'due_date': (datetime.now() + timedelta(days=3)).strftime('%Y-%m-%d'),
            'status': 'Not Started',
            'progress': 0,
            'description': 'Create wireframes and high-fidelity mockups'
        },
        {
            'name': 'Write product documentation',
            'category': 'Admin',
            'assigned_to': 'Carol Davis',
            'priority': 'Low',
            'due_date': (datetime.now() + timedelta(days=10)).strftime('%Y-%m-%d'),
            'status': 'Not Started',
            'progress': 0,
            'description': 'Document API endpoints and usage'
        },
        {
            'name': 'Plan Q1 marketing campaign',
            'category': 'Marketing',
            'assigned_to': 'David Lee',
            'priority': 'Critical',
            'due_date': (datetime.now() + timedelta(days=2)).strftime('%Y-%m-%d'),
            'status': 'In Progress',
            'progress': 60,
            'description': 'Prepare marketing strategy for Q1'
        },
        {
            'name': 'Customer support training',
            'category': 'Support',
            'assigned_to': 'Emma Wilson',
            'priority': 'Medium',
            'due_date': (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d'),
            'status': 'Not Started',
            'progress': 0,
            'description': 'Train new support team members'
        }
    ]
    
    print("\n📝 Adding sample tasks...")
    for task in sample_tasks:
        tms.add_task(task)
    
    print("\n" + "=" * 60)
    print("✅ TASK MANAGEMENT SYSTEM CREATED!")
    print("=" * 60)
    print(f"\n📁 File: task_management.xlsx")
    print("\n📊 Sheets created:")
    print("   1. Dashboard - Live overview with KPIs")
    print("   2. Task Entry - Add and manage all tasks")
    print("   3-10. Category Sheets - Auto-sorted tasks")
    print("   11. Team View - Tasks by team member")
    print("   12. Analytics - Statistics and reports")
    print("   13. Archive - Completed tasks")
    print("\n🎯 Features:")
    print("   ✓ Automatic task routing by category")
    print("   ✓ Live dashboard with real-time KPIs")
    print("   ✓ Outstanding tasks tracking")
    print("   ✓ Team member assignment tracking")
    print("   ✓ Priority and status color coding")
    print("   ✓ Data validation dropdowns")
    print("   ✓ Progress tracking")
    print("   ✓ Analytics and reporting")
    print("\n🔄 Next: Open the file and start adding tasks!")
    print("=" * 60 + "\n")
