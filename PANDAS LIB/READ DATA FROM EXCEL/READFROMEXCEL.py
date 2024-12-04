import pandas as pd

file_path = 'employees.xlsx'
df = pd.read_excel(file_path)

print("DATA OF FIRST 5 ROW :")
print(df.head())

department_counts = df['Department'].value_counts()
print("\nNumber of employees in each department:")
print(department_counts)


average_salary = df['Salary'].mean()
print(f"\n Average salaries in the company : {average_salary}")

max_salary = df['Salary'].max()
min_salary = df['Salary'].min()
highest_paid_employee = df[df['Salary'] == max_salary]
lowest_paid_employee = df[df['Salary'] == min_salary]

print(f"\nHighest salary in the company: {max_salary}, employee: {highest_paid_employee['Name'].values[0]}")
print(f"Smallest salary in the company: {min_salary}, employee: {lowest_paid_employee['Name'].values[0]}")


df['Joining_Date'] = pd.to_datetime(df['Joining_Date'])
oldest_employee = df[df['Joining_Date'] == df['Joining_Date'].min()]
print(f"\nThe most senior employee in the company: {oldest_employee['Name'].values[0]}")

salary_threshold = 60000  # الحد الأدنى لمتوسط الراتب
departments_with_high_salary = df.groupby('Department')['Salary'].mean()
high_salary_departments = departments_with_high_salary[departments_with_high_salary > salary_threshold]
print("\nالأقسام التي يتجاوز متوسط الرواتب فيها 60000:")
print(high_salary_departments)

# **حفظ النتائج في ملف Excel جديد**
output_data = {
    'Department Counts': department_counts,
    'High Salary Departments': high_salary_departments
}

# تحويل النتائج إلى DataFrame
results_df = pd.DataFrame(output_data)

# حفظ النتائج
results_df.to_excel('company_analysis.xlsx', index=True, sheet_name='Analysis')

print("\nتم حفظ التحليلات في ملف company_analysis.xlsx")
