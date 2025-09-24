import os
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import random

def create_sales_data():
    """
    Create a sample sales dataset with various data types.
    """
    # Set random seed for reproducibility
    np.random.seed(42)
    
    # Create date range for the past year
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365)
    dates = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # Product categories and regions
    products = ['Laptop', 'Smartphone', 'Tablet', 'Monitor', 'Keyboard', 'Mouse', 'Headphones', 'Printer']
    categories = ['Electronics', 'Accessories', 'Peripherals']
    regions = ['North', 'South', 'East', 'West', 'Central']
    
    # Create random data
    n_samples = len(dates)
    
    data = {
        'Date': np.random.choice(dates, n_samples),
        'Product': np.random.choice(products, n_samples),
        'Category': [random.choice(categories) for _ in range(n_samples)],
        'Region': np.random.choice(regions, n_samples),
        'Units_Sold': np.random.randint(1, 50, n_samples),
        'Unit_Price': np.random.uniform(10, 1000, n_samples).round(2),
        'Shipping_Cost': np.random.uniform(5, 50, n_samples).round(2),
        'Discount_Pct': np.random.choice([0, 5, 10, 15, 20], n_samples),
        'Customer_Rating': np.random.randint(1, 6, n_samples),
        'Returns': np.random.binomial(1, 0.05, n_samples)  # 5% return rate
    }
    
    # Calculate total sales
    data['Total_Sales'] = (data['Units_Sold'] * data['Unit_Price'] * (1 - data['Discount_Pct']/100)).round(2)
    
    # Add some missing values
    for col in ['Customer_Rating', 'Shipping_Cost', 'Discount_Pct']:
        mask = np.random.random(n_samples) < 0.05  # 5% missing rate
        data[col] = pd.Series(data[col])
        data[col][mask] = np.nan
    
    # Add some outliers
    outlier_indices = np.random.choice(range(n_samples), size=5, replace=False)
    data['Units_Sold'][outlier_indices] = np.random.randint(100, 200, 5)
    data['Unit_Price'][outlier_indices] = np.random.uniform(1500, 2000, 5).round(2)
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Add a few more calculated columns
    df['Month'] = df['Date'].dt.month_name()
    df['Day_of_Week'] = df['Date'].dt.day_name()
    df['Is_Weekend'] = df['Date'].dt.dayofweek >= 5
    df['Profit'] = (df['Total_Sales'] - df['Shipping_Cost']).round(2)
    df['Profit_Margin'] = ((df['Profit'] / df['Total_Sales']) * 100).round(2)
    
    return df

def create_employee_data():
    """
    Create a sample employee dataset.
    """
    # Set random seed for reproducibility
    np.random.seed(43)
    
    # Number of employees
    n_employees = 100
    
    # Create random data
    departments = ['HR', 'IT', 'Finance', 'Marketing', 'Sales', 'Operations', 'R&D']
    job_titles = ['Manager', 'Director', 'Associate', 'Analyst', 'Specialist', 'Coordinator', 'Assistant']
    education = ['High School', 'Bachelor', 'Master', 'PhD']
    
    # Generate hire dates over the past 10 years
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365*10)
    hire_dates = pd.date_range(start=start_date, end=end_date, periods=n_employees)
    
    data = {
        'Employee_ID': range(1001, 1001 + n_employees),
        'Department': np.random.choice(departments, n_employees),
        'Job_Title': [f"{random.choice(job_titles)} {random.choice(['I', 'II', 'III', 'IV', 'V'])}" for _ in range(n_employees)],
        'Education': np.random.choice(education, n_employees),
        'Age': np.random.randint(22, 65, n_employees),
        'Years_Experience': np.random.randint(0, 30, n_employees),
        'Salary': np.random.normal(60000, 20000, n_employees).round(2),
        'Hire_Date': hire_dates,
        'Performance_Score': np.random.uniform(1, 5, n_employees).round(1),
        'Satisfaction_Score': np.random.uniform(1, 10, n_employees).round(1),
        'Training_Hours': np.random.randint(0, 100, n_employees),
        'Promotion_Eligible': np.random.choice([True, False], n_employees)
    }
    
    # Add some missing values
    for col in ['Performance_Score', 'Satisfaction_Score', 'Training_Hours']:
        mask = np.random.random(n_employees) < 0.05  # 5% missing rate
        data[col] = pd.Series(data[col])
        data[col][mask] = np.nan
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Add a few more calculated columns
    df['Years_at_Company'] = ((datetime.now() - df['Hire_Date']).dt.days / 365).round(1)
    df['Salary_per_Experience'] = (df['Salary'] / (df['Years_Experience'] + 1)).round(2)
    
    return df

def create_customer_data():
    """
    Create a sample customer dataset.
    """
    # Set random seed for reproducibility
    np.random.seed(44)
    
    # Number of customers
    n_customers = 200
    
    # Create random data
    genders = ['Male', 'Female', 'Non-binary', 'Prefer not to say']
    segments = ['Premium', 'Standard', 'Basic']
    acquisition_channels = ['Organic Search', 'Paid Search', 'Social Media', 'Email', 'Referral', 'Direct']
    
    # Generate registration dates over the past 5 years
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365*5)
    reg_dates = pd.date_range(start=start_date, end=end_date, periods=n_customers)
    
    data = {
        'Customer_ID': range(5001, 5001 + n_customers),
        'Age': np.random.randint(18, 80, n_customers),
        'Gender': np.random.choice(genders, n_customers),
        'Location': [f"City-{i}" for i in np.random.randint(1, 31, n_customers)],
        'Segment': np.random.choice(segments, n_customers),
        'Registration_Date': reg_dates,
        'Acquisition_Channel': np.random.choice(acquisition_channels, n_customers),
        'Total_Purchases': np.random.randint(0, 50, n_customers),
        'Total_Spend': np.random.exponential(scale=500, size=n_customers).round(2),
        'Average_Order_Value': np.random.normal(100, 30, n_customers).round(2),
        'Days_Since_Last_Purchase': np.random.randint(1, 365, n_customers),
        'Is_Active': np.random.choice([True, False], n_customers, p=[0.8, 0.2])
    }
    
    # Add some missing values
    for col in ['Age', 'Total_Spend', 'Days_Since_Last_Purchase']:
        mask = np.random.random(n_customers) < 0.05  # 5% missing rate
        data[col] = pd.Series(data[col])
        data[col][mask] = np.nan
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Add a few more calculated columns
    df['Customer_Tenure_Days'] = (datetime.now() - df['Registration_Date']).dt.days
    df['Purchase_Frequency'] = (df['Total_Purchases'] / (df['Customer_Tenure_Days'] / 30)).round(2)
    df['Customer_Value'] = (df['Total_Spend'] * df['Purchase_Frequency'] / 12).round(2)
    
    return df

def main():
    """
    Main function to create sample Excel files.
    """
    # Create output directory if it doesn't exist
    samples_dir = 'samples'
    if not os.path.exists(samples_dir):
        os.makedirs(samples_dir)
    
    # Create sales data
    print("Creating sales data...")
    sales_df = create_sales_data()
    sales_file = os.path.join(samples_dir, 'sales_data.xlsx')
    sales_df.to_excel(sales_file, index=False)
    print(f"Sales data saved to {sales_file}")
    
    # Create employee data
    print("Creating employee data...")
    employee_df = create_employee_data()
    employee_file = os.path.join(samples_dir, 'employee_data.xlsx')
    employee_df.to_excel(employee_file, index=False)
    print(f"Employee data saved to {employee_file}")
    
    # Create customer data
    print("Creating customer data...")
    customer_df = create_customer_data()
    customer_file = os.path.join(samples_dir, 'customer_data.xlsx')
    customer_df.to_excel(customer_file, index=False)
    print(f"Customer data saved to {customer_file}")
    
    # Create a multi-sheet Excel file
    print("Creating multi-sheet Excel file...")
    with pd.ExcelWriter(os.path.join(samples_dir, 'sample_data.xlsx')) as writer:
        sales_df.to_excel(writer, sheet_name='Sales', index=False)
        employee_df.to_excel(writer, sheet_name='Employees', index=False)
        customer_df.to_excel(writer, sheet_name='Customers', index=False)
    print(f"Multi-sheet data saved to {os.path.join(samples_dir, 'sample_data.xlsx')}")
    
    print("Sample data creation complete!")

if __name__ == "__main__":
    main()