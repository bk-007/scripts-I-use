import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os

class EmployeeManagementSystemGenerator:
    def __init__(self, file_path="Employee_Management_System.xlsx"):
        self.file_path = file_path
        self.generate_comprehensive_system()
    
    def generate_comprehensive_system(self):
        """Generate the complete employee management system with sample data"""
        print("Generating Comprehensive Employee Management System...")
        
        # Generate all data tables
        df_employees = self.create_employees_table()
        df_training_processes = self.create_training_processes_table()
        df_training_status = self.create_training_status_table()
        df_one_on_ones = self.create_one_on_ones_table()
        df_projects = self.create_projects_table()
        df_onboarding_tasks = self.create_onboarding_tasks_table()
        df_lookup_values = self.create_lookup_values_table()
        
        # Write all sheets to Excel
        with pd.ExcelWriter(self.file_path, engine='openpyxl') as writer:
            df_employees.to_excel(writer, sheet_name='Employees', index=False)
            df_training_processes.to_excel(writer, sheet_name='Training_Processes', index=False)
            df_training_status.to_excel(writer, sheet_name='Training_Status', index=False)
            df_one_on_ones.to_excel(writer, sheet_name='One_on_Ones', index=False)
            df_projects.to_excel(writer, sheet_name='Projects', index=False)
            df_onboarding_tasks.to_excel(writer, sheet_name='Onboarding_Tasks', index=False)
            df_lookup_values.to_excel(writer, sheet_name='Lookup_Values', index=False)
        
        print(f"✅ Complete system generated: {self.file_path}")
        self.print_summary()
    
    def create_employees_table(self):
        """Create comprehensive employees table"""
        departments = ['Operations', 'Finance', 'Sales', 'Legal', 'HR', 'IT', 'Marketing']
        positions = ['Analyst', 'Specialist', 'Manager', 'Director', 'Coordinator', 'Associate']
        
        employees_data = {
            'Employee_ID': [f'EMP{str(i).zfill(3)}' for i in range(1, 21)],
            'Employee_Name': [
                'John Smith', 'Sarah Johnson', 'Mike Brown', 'Lisa Davis', 'Tom Wilson',
                'Anna Garcia', 'David Lee', 'Emma Martinez', 'Chris Taylor', 'Maya Patel',
                'James Anderson', 'Jessica White', 'Robert Clark', 'Amy Rodriguez', 'Kevin Chen',
                'Nicole Thompson', 'Daniel Kim', 'Olivia Jackson', 'Ryan Murphy', 'Sophia Liu'
            ],
            'Email': [f'{name.lower().replace(" ", ".")}@company.com' for name in [
                'John Smith', 'Sarah Johnson', 'Mike Brown', 'Lisa Davis', 'Tom Wilson',
                'Anna Garcia', 'David Lee', 'Emma Martinez', 'Chris Taylor', 'Maya Patel',
                'James Anderson', 'Jessica White', 'Robert Clark', 'Amy Rodriguez', 'Kevin Chen',
                'Nicole Thompson', 'Daniel Kim', 'Olivia Jackson', 'Ryan Murphy', 'Sophia Liu'
            ]],
            'Department': [random.choice(departments) for _ in range(20)],
            'Position': [random.choice(positions) for _ in range(20)],
            'Hire_Date': [datetime.now() - timedelta(days=random.randint(30, 1095)) for _ in range(20)],
            'Manager': [
                'Jane Doe', 'Robert Clark', 'Amy White', 'David Lee', 'Jane Doe',
                'Robert Clark', 'Amy White', 'Jane Doe', 'David Lee', 'Amy White',
                'Jane Doe', 'Robert Clark', 'David Lee', 'Amy White', 'Jane Doe',
                'Robert Clark', 'Amy White', 'Jane Doe', 'David Lee', 'Amy White'
            ],
            'Employment_Status': ['Active'] * 20,
            'Onboarding_Status': [random.choice(['Completed', 'In Progress', 'Not Started']) for _ in range(20)],
            'Last_One_on_One_Date': [datetime.now() - timedelta(days=random.randint(1, 90)) for _ in range(20)],
            'Performance_Rating': [random.choice(['Exceeds', 'Meets', 'Below', 'New Employee']) for _ in range(20)],
            'Location': [random.choice(['Remote', 'Office', 'Hybrid']) for _ in range(20)]
        }
        
        return pd.DataFrame(employees_data)
    
    def create_training_processes_table(self):
        """Create comprehensive training processes table"""
        training_data = {
            'Process_ID': [f'PROC{str(i).zfill(3)}' for i in range(1, 21)],
            'Process_Name': [
                'New Employee Orientation', 'Company Policies & Procedures', 'Safety Training',
                'Invoice Processing', 'Quality Control', 'Customer Onboarding',
                'Compliance Review', 'Sales Forecasting', 'Risk Assessment',
                'Data Analysis', 'Project Management', 'Leadership Development',
                'Communication Skills', 'Time Management', 'Conflict Resolution',
                'Technical Writing', 'Presentation Skills', 'Team Building',
                'Performance Management', 'Strategic Planning'
            ],
            'Department': [
                'HR', 'HR', 'Operations', 'Finance', 'Operations', 'Sales',
                'Legal', 'Sales', 'Finance', 'IT', 'Operations', 'HR',
                'HR', 'HR', 'HR', 'Marketing', 'Sales', 'HR',
                'HR', 'Operations'
            ],
            'Training_Type': [
                'Onboarding', 'Onboarding', 'Compliance', 'Skill Development', 'Compliance', 'Skill Development',
                'Compliance', 'Skill Development', 'Compliance', 'Skill Development', 'Skill Development', 'Leadership',
                'Soft Skills', 'Soft Skills', 'Soft Skills', 'Skill Development', 'Soft Skills', 'Soft Skills',
                'Leadership', 'Skill Development'
            ],
            'Estimated_Hours': [8, 4, 6, 12, 16, 8, 20, 10, 14, 24, 32, 16, 8, 6, 10, 12, 8, 16, 20, 24],
            'Difficulty_Level': [
                'Beginner', 'Beginner', 'Beginner', 'Intermediate', 'Intermediate', 'Beginner',
                'Advanced', 'Intermediate', 'Advanced', 'Intermediate', 'Advanced', 'Advanced',
                'Beginner', 'Beginner', 'Intermediate', 'Intermediate', 'Beginner', 'Intermediate',
                'Advanced', 'Advanced'
            ],
            'Is_Required': [
                True, True, True, False, True, False, True, False, True, False,
                False, False, False, False, False, False, False, False, False, False
            ],
            'Frequency': [
                'Once', 'Once', 'Annual', 'Once', 'Annual', 'Once', 'Annual', 'Once', 'Annual', 'Once',
                'Once', 'Once', 'Once', 'Once', 'Once', 'Once', 'Once', 'Once', 'Once', 'Once'
            ]
        }
        
        return pd.DataFrame(training_data)
    
    def create_training_status_table(self):
        """Create training status tracking table"""
        # Generate 50 training records across employees
        training_records = []
        
        for i in range(1, 51):
            employee_id = f'EMP{str(random.randint(1, 20)).zfill(3)}'
            process_id = f'PROC{str(random.randint(1, 20)).zfill(3)}'
            status = random.choice(['Completed', 'In Progress', 'Planned', 'On Hold'])
            
            start_date = None
            completion_date = None
            planned_completion = datetime.now() + timedelta(days=random.randint(1, 180))
            
            if status in ['Completed', 'In Progress']:
                start_date = datetime.now() - timedelta(days=random.randint(1, 90))
                if status == 'Completed':
                    completion_date = start_date + timedelta(days=random.randint(1, 30))
                    planned_completion = completion_date
            
            progress = 100 if status == 'Completed' else random.randint(0, 90) if status == 'In Progress' else 0
            
            training_records.append({
                'Record_ID': i,
                'Employee_ID': employee_id,
                'Process_ID': process_id,
                'Training_Status': status,
                'Start_Date': start_date,
                'Completion_Date': completion_date,
                'Planned_Completion': planned_completion,
                'Progress_Percentage': progress,
                'Assigned_By': random.choice(['Jane Doe', 'Robert Clark', 'Amy White', 'David Lee']),
                'Priority': random.choice(['High', 'Medium', 'Low']),
                'Notes': random.choice([
                    'Excellent performance', 'On track', 'Needs support', 'Quick learner',
                    'Requires additional time', 'Meeting expectations', 'Outstanding progress',
                    'Slow start but improving', 'Complex material', 'Exceeding expectations'
                ])
            })
        
        return pd.DataFrame(training_records)
    
    def create_one_on_ones_table(self):
        """Create one-on-one meetings tracking table"""
        meeting_types = ['Weekly Check-in', 'Monthly Review', 'Quarterly Review', 'Goal Setting', 'Performance Review']
        
        one_on_one_records = []
        
        for i in range(1, 31):
            employee_id = f'EMP{str(random.randint(1, 20)).zfill(3)}'
            meeting_date = datetime.now() - timedelta(days=random.randint(1, 180))
            
            one_on_one_records.append({
                'Meeting_ID': f'MEET{str(i).zfill(3)}',
                'Employee_ID': employee_id,
                'Manager_Name': random.choice(['Jane Doe', 'Robert Clark', 'Amy White', 'David Lee']),
                'Meeting_Date': meeting_date,
                'Meeting_Type': random.choice(meeting_types),
                'Duration_Minutes': random.choice([30, 45, 60]),
                'Goals_Discussed': random.choice([
                    'Career development goals', 'Project objectives', 'Skill improvement',
                    'Work-life balance', 'Team collaboration', 'Process improvements'
                ]),
                'Challenges_Raised': random.choice([
                    'Time management', 'Resource constraints', 'Technical difficulties',
                    'Communication barriers', 'Workload concerns', 'Training gaps'
                ]),
                'Action_Items': random.choice([
                    'Complete training module', 'Schedule follow-up meeting', 'Research new tools',
                    'Join project team', 'Attend workshop', 'Prepare presentation'
                ]),
                'Employee_Satisfaction': random.randint(7, 10),
                'Next_Meeting_Date': meeting_date + timedelta(days=random.randint(7, 30)),
                'Meeting_Status': random.choice(['Completed', 'Scheduled', 'Cancelled']),
                'Notes': 'Regular check-in meeting to discuss progress and challenges.'
            })
        
        return pd.DataFrame(one_on_one_records)
    
    def create_projects_table(self):
        """Create project tracking table"""
        project_names = [
            'Customer Portal Upgrade', 'Process Automation', 'Data Migration', 'Mobile App Development',
            'Security Audit', 'Website Redesign', 'CRM Implementation', 'Training Platform',
            'Quality Management System', 'Employee Onboarding Portal', 'Analytics Dashboard',
            'Inventory Management', 'Customer Feedback System', 'Document Management',
            'Performance Tracking Tool'
        ]
        
        project_records = []
        
        for i, project_name in enumerate(project_names, 1):
            start_date = datetime.now() - timedelta(days=random.randint(30, 365))
            due_date = start_date + timedelta(days=random.randint(30, 180))
            
            project_records.append({
                'Project_ID': f'PROJ{str(i).zfill(3)}',
                'Project_Name': project_name,
                'Employee_ID': f'EMP{str(random.randint(1, 20)).zfill(3)}',
                'Project_Manager': random.choice(['Jane Doe', 'Robert Clark', 'Amy White', 'David Lee']),
                'Status': random.choice(['Planning', 'In Progress', 'On Hold', 'Completed', 'Cancelled']),
                'Priority': random.choice(['High', 'Medium', 'Low']),
                'Start_Date': start_date,
                'Due_Date': due_date,
                'Progress_Percentage': random.randint(0, 100),
                'Budget': random.randint(5000, 100000),
                'Department': random.choice(['IT', 'Operations', 'Finance', 'HR', 'Marketing']),
                'Last_Update': datetime.now() - timedelta(days=random.randint(1, 30)),
                'Risk_Level': random.choice(['Low', 'Medium', 'High']),
                'Team_Size': random.randint(2, 8),
                'Notes': 'Project progressing according to schedule.'
            })
        
        return pd.DataFrame(project_records)
    
    def create_onboarding_tasks_table(self):
        """Create onboarding tasks tracking table"""
        onboarding_tasks = [
            'Complete paperwork', 'IT setup and accounts', 'Office tour', 'Meet team members',
            'Review company handbook', 'Safety training', 'Department orientation', 'Assign mentor',
            'Set up workspace', 'First week check-in', 'Complete required training',
            'Review job description', 'Set initial goals', '30-day review', '60-day review', '90-day review'
        ]
        
        onboarding_records = []
        task_id = 1
        
        # Create tasks for employees with onboarding status
        for emp_num in range(1, 21):
            employee_id = f'EMP{str(emp_num).zfill(3)}'
            
            # Assign 8-12 random tasks per employee
            num_tasks = random.randint(8, 12)
            selected_tasks = random.sample(onboarding_tasks, num_tasks)
            
            for task_name in selected_tasks:
                due_date = datetime.now() + timedelta(days=random.randint(-30, 60))
                
                onboarding_records.append({
                    'Task_ID': f'TASK{str(task_id).zfill(3)}',
                    'Employee_ID': employee_id,
                    'Task_Name': task_name,
                    'Department': random.choice(['HR', 'IT', 'Operations', 'All']),
                    'Due_Date': due_date,
                    'Status': random.choice(['Completed', 'In Progress', 'Pending', 'Overdue']),
                    'Assigned_To': random.choice(['HR Team', 'IT Team', 'Manager', 'Mentor']),
                    'Completion_Date': due_date - timedelta(days=random.randint(1, 5)) if random.choice([True, False]) else None,
                    'Priority': random.choice(['High', 'Medium', 'Low']),
                    'Estimated_Hours': random.choice([0.5, 1, 2, 4, 8]),
                    'Notes': random.choice([
                        'Standard onboarding task', 'Critical for first week', 'Department specific',
                        'Requires manager approval', 'Self-paced learning', 'Scheduled session'
                    ])
                })
                task_id += 1
        
        return pd.DataFrame(onboarding_records)
    
    def create_lookup_values_table(self):
        """Create lookup values for form dropdowns"""
        lookup_data = {
            'Category': [
                'Department', 'Department', 'Department', 'Department', 'Department', 'Department', 'Department',
                'Training_Status', 'Training_Status', 'Training_Status', 'Training_Status',
                'Project_Status', 'Project_Status', 'Project_Status', 'Project_Status', 'Project_Status',
                'Priority', 'Priority', 'Priority',
                'Performance_Rating', 'Performance_Rating', 'Performance_Rating', 'Performance_Rating',
                'Meeting_Type', 'Meeting_Type', 'Meeting_Type', 'Meeting_Type', 'Meeting_Type'
            ],
            'Value': [
                'Operations', 'Finance', 'Sales', 'Legal', 'HR', 'IT', 'Marketing',
                'Completed', 'In Progress', 'Planned', 'On Hold',
                'Planning', 'In Progress', 'On Hold', 'Completed', 'Cancelled',
                'High', 'Medium', 'Low',
                'Exceeds', 'Meets', 'Below', 'New Employee',
                'Weekly Check-in', 'Monthly Review', 'Quarterly Review', 'Goal Setting', 'Performance Review'
            ],
            'Display_Order': [
                1, 2, 3, 4, 5, 6, 7,
                1, 2, 3, 4,
                1, 2, 3, 4, 5,
                1, 2, 3,
                1, 2, 3, 4,
                1, 2, 3, 4, 5
            ]
        }
        
        return pd.DataFrame(lookup_data)
    
    def print_summary(self):
        """Print system summary"""
        print("\n" + "="*60)
        print("📊 EMPLOYEE MANAGEMENT SYSTEM SUMMARY")
        print("="*60)
        print("📁 Excel Sheets Created:")
        print("   • Employees (20 sample employees)")
        print("   • Training_Processes (20 training programs)")
        print("   • Training_Status (50 training records)")
        print("   • One_on_Ones (30 meeting records)")
        print("   • Projects (15 active projects)")
        print("   • Onboarding_Tasks (180+ onboarding tasks)")
        print("   • Lookup_Values (dropdown options)")
        print("\n🔧 Ready for:")
        print("   • Upload to Excel Online")
        print("   • Power Automate integration")
        print("   • Power BI dashboard creation")
        print("   • Microsoft Forms connection")
        print("="*60)

# Usage
if __name__ == "__main__":
    print("🚀 Generating Comprehensive Employee Management System...")
    system = EmployeeManagementSystemGenerator()
    print("\n✅ System ready for deployment!")