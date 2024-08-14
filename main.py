import pandas as pd 
import numpy as np
file_path="C:\\Users\\pc\\Desktop\\python projects\\task\\Employee Sample Data - Copy.xlsx"
df=pd.read_excel(file_path)
print(df)
print(df.info())
columns_to_check=[col for col in df.columns if col !='Exit Date']
df=df.dropna(subset=columns_to_check)


new = {
    'EEID': [101, 102, 103, 104, 105],
    'Full Name': ['Osama', 'Ibrahim', 'Moumen', 'Lujain', 'Radi'],
    'Job Title': ['Manager', 'Sales Representative', 'HR Manager', 'Software Engineer', 'Marketing Engineer'],
    'Department': ['Engineering', 'Sales', 'HR', 'IT', 'Marketing'],
    'Business Unit':['Manufacturing','Manufacturing','Manufacturing','Manufacturing','Manufacturing'],
    'Gender':['Male','Male','Male','Female','Male'],
    'Ethnicity':['African American ','Arabian','Arabian','Arabian','Arabian'],
    'Age': [22,24 , 28, 23, 23],
    'Hire Date': [pd.Timestamp('2018-10-15'), pd.Timestamp('2019-07-17'), pd.Timestamp('2015-12-14'), pd.Timestamp('2019-04-11'),pd.Timestamp('2017-05-10')],
    'Annual Salary':[113527,100000,55000,66000,77000],
    'Bonus %':[20,10,15,5,8],
    'Country':['Tanzania','Thailand','Syria','Jordan','Jordan'],
    'City':['Zanzibar','Bangkok','Daraa','Amman','Irbid'],
    'Exit Date': ['until now', pd.Timestamp('2024-01-15'), 'until now', pd.Timestamp('2024-02-10'), 'until now']
}

df.iloc[:5] = pd.DataFrame(new)

print(df.head(10))




max_salary_row = df.loc[df['Annual Salary'].idxmax()]
print('The employee with the largest salary\n',max_salary_row)


department_group = df.groupby('Department').agg({
    'Age': 'mean',
    'Annual Salary': 'mean'
})

department_group = department_group.rename(columns={'Age': 'Average Age', 'Annual Salary': 'Average Salary'})
print("Average age and average salary by department:")
print(department_group)


dept_eth_group = df.groupby(['Department', 'Ethnicity']).agg({
    'Age': ['max', 'min'],
    'Annual Salary': 'median'
})

dept_eth_group.columns = ['Max Age', 'Min Age', 'Median Salary']

print("\nMax age, min age, and median salary by department and ethnicity:")
print(dept_eth_group)

with pd.ExcelWriter('cleanedEmployeeSampleData.xlsx') as writer:
  
    df.to_excel(writer, sheet_name='Cleaned Data')
    max_salary_row.to_excel(writer, sheet_name='Max Salary Employee')
    department_group.to_excel(writer, sheet_name='Avg Age and Salary by Dept')
    dept_eth_group.to_excel(writer, sheet_name='Age and Salary by Dept+Ethn')
    
print("Data has been saved to 'cleanedEmployeeSampleData.xlsx'")