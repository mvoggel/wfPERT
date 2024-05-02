import pandas as pd
import numpy as np

# Read the Excel file into a DataFrame
def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to filter DataFrame based on user input
def filter_data(df, filter_criteria, max_job_size):
    # Define the list of columns to be filtered
    columns_to_filter = ['Business Category', 'Owner: Name', 'Job Size']

    # Separate the filter criteria for each column
    business_category = filter_criteria.get('Business Category', '')
    owner_name = filter_criteria.get('Owner: Name', '')

    # Apply filtering operations based on user input
    filtered_df = df[df['Business Category'].str.contains(business_category, na=False) &
                     df['Owner: Name'].str.contains(owner_name, na=False) &
                     (df['Job Size'] <= max_job_size)]

    return filtered_df

# Function to calculate statistics for Actual Duration (days), Actual Hours, and Expected Time
def calculate_statistics(filtered_df):
    duration_stats = {
        'Highest Actual Duration (days)': filtered_df['Actual Duration (days)'].max(),
        'Lowest Actual Duration (days)': filtered_df['Actual Duration (days)'].min(),
        'Average Actual Duration (days)': filtered_df['Actual Duration (days)'].mean(),
        'Total Sum of Actual Duration (days)': filtered_df['Actual Duration (days)'].sum()
    }

    hours_stats = {
        'Highest Actual Hours': filtered_df['Actual Hours'].max(),
        'Lowest Actual Hours': filtered_df['Actual Hours'].min(),
        'Average Actual Hours': filtered_df['Actual Hours'].mean(),
        'Total Sum of Actual Hours': filtered_df['Actual Hours'].sum()
    }

    # Calculate Expected Time
    expected_time_days = (filtered_df['Actual Duration (days)'].min() + (filtered_df['Actual Duration (days)'].mean() * 4) + filtered_df['Actual Duration (days)'].max()) / 6
    expected_time_hours = (filtered_df['Actual Hours'].min() + (filtered_df['Actual Hours'].mean() * 4) + filtered_df['Actual Hours'].max()) / 6

    return duration_stats, hours_stats, expected_time_days, expected_time_hours

# Function to calculate variance and standard deviation
def calculate_variance_and_std_dev(duration_values, hours_values):
    duration_variance = pd.Series(duration_values).var()
    duration_std_dev = pd.Series(duration_values).std()

    hours_variance = pd.Series(hours_values).var()
    hours_std_dev = pd.Series(hours_values).std()

    return duration_variance, duration_std_dev, hours_variance, hours_std_dev

# Function to calculate 95% Confidence Stat
def calculate_95_confidence_stat(variance):
    return 1.6 * np.sqrt(variance)

# Function to print the results
def print_results(duration_stats, hours_stats, expected_time_hours, expected_time_days, duration_variance, duration_std_dev, hours_variance, hours_std_dev, duration_95_confidence_stat, hours_95_confidence_stat):
    print("Actual Duration:")
    for key, value in duration_stats.items():
        print(f"{key}: {value}")
    
    print("\nActual Hours:")
    for key, value in hours_stats.items():
        print(f"{key}: {value}")

    print("\nExpected Time:")
    print(f"Expected Time Days: {expected_time_days}")
    print(f"Expected Time Hours: {expected_time_hours}")

    print("\nVariance and Standard Deviation:")
    print(f"Variance of Actual Duration: {duration_variance}")
    print(f"Standard Deviation of Actual Duration: {duration_std_dev}")
    print(f"Variance of Actual Hours: {hours_variance}")
    print(f"Standard Deviation of Actual Hours: {hours_std_dev}")

    print("\n95% Confidence Stat:")
    print(f"95% Confidence Stat - Days: {expected_time_days + duration_95_confidence_stat}")
    print(f"95% Confidence Stat - Hours: {expected_time_hours + hours_95_confidence_stat}")

# Main function
def main():
    # Path to Excel file on desktop
    file_path = "/Users/matthewvoggel/Desktop/Pert Test.xlsx"  # Update "YourUsername" with your actual username

    # Read Excel file from desktop
    df = read_excel_file(file_path)

    # Get user input for filtering criteria
    filter_criteria = {
        'Business Category': input("Enter Business Category: "),
        'Owner: Name': input("Enter Owner's Name: ")
    }

    # Get user input for maximum Job Size
    max_job_size = float(input("Enter maximum Job Size: "))

    # Filter the data based on user input
    filtered_df = filter_data(df, filter_criteria, max_job_size)

    # Calculate statistics for Actual Duration (days), Actual Hours, and Expected Time
    duration_stats, hours_stats, expected_time_days, expected_time_hours = calculate_statistics(filtered_df)

    # Print the statistics
    duration_values = filtered_df['Actual Duration (days)'].tolist()
    hours_values = filtered_df['Actual Hours'].tolist()

    duration_variance, duration_std_dev, hours_variance, hours_std_dev = calculate_variance_and_std_dev(duration_values, hours_values)

    duration_95_confidence_stat = calculate_95_confidence_stat(duration_variance)
    hours_95_confidence_stat = calculate_95_confidence_stat(hours_variance)

    print_results(duration_stats, hours_stats, expected_time_hours, expected_time_days, duration_variance, duration_std_dev, hours_variance, hours_std_dev, duration_95_confidence_stat, hours_95_confidence_stat)

if __name__ == "__main__":
    main()
