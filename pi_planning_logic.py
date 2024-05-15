import pandas as pd

def filter_excel(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)
    
    # Filter the DataFrame based on the specified columns
    filtered_df = df[[
        "Is this a new, urgent request that was not planned for the Planning Interval/PI that it is being completed in?",
        "What is the original planned PI that this project should be completed in?",
        "What is the NEW planned PI that this project should be completed in?",
        "What is the original planned ITERATION that this project should be completed in?",
        "What is the NEW planned ITERATION that this project should be completed in?",
        "Job Size",
        "WSJF"
    ]]
    
    return filtered_df


def analyze_pi_planning(filtered_df, file_path):
    # Total number of rows in the DataFrame
    total_rows = len(filtered_df)

    # Total number of 'Yes' and 'No' answers for urgent requests
    total_yes = (filtered_df['Is this a new, urgent request that was not planned for the Planning Interval/PI that it is being completed in?'] == 'Yes').sum()
    total_no = (filtered_df['Is this a new, urgent request that was not planned for the Planning Interval/PI that it is being completed in?'] == 'No').sum()

    # Total number of rows with either 'Yes' or 'No' answer
    total_yes_no = total_yes + total_no

    # Total number of rows with blank cells in the specified column
    total_blank_cells = filtered_df['Is this a new, urgent request that was not planned for the Planning Interval/PI that it is being completed in?'].isnull().sum()

    # Calculate average job size
    total_job_size = filtered_df['Job Size'].sum()
    average_job_size = total_job_size / total_rows

    # Count projects by each job size
    job_size_counts = filtered_df['Job Size'].value_counts()

    # Calculate average WSJF
    total_wsjf = filtered_df['WSJF'].sum()
    average_wsjf = total_wsjf / total_rows

    # Define WSJF ranges
    wsjf_ranges = [
        (0.0, 0.49),
        (0.5, 0.99),
        (1.0, 1.99),
        (2.0, 3.49),
        (3.5, 5.0),
        (5.1, 8.49),
        (8.5, 11.99),
        (12.0, 20.99),
        (21.0, float('inf'))
    ]

    # Count projects by WSJF ranges
    wsjf_range_counts = {}
    for wsjf_range in wsjf_ranges:
        lower, upper = wsjf_range
        count = ((filtered_df['WSJF'] >= lower) & (filtered_df['WSJF'] <= upper)).sum()
        wsjf_range_counts[f'{lower} - {upper}'] = count
        
        

    # Print the calculated information
    print("\nSummary for Urgent requests not part of PI planning:")
    print(f"\tTotal number of projects: {total_rows}")
    print(f"\tTotal number of Urgent requests, marked 'Yes': {total_yes}")
    print(f"\tTotal number of non-urgent requests, marked 'No': {total_no}")
    print(f"\tTotal number of rows with either 'Yes' or 'No' answer: {total_yes_no}")
    print(f"\tTotal number of projects with with this field blank: {total_blank_cells}")

    print("\nSummary of PI-Planned Projects info.:")
    print(f"\tTotal average Job Size of all projects: {average_job_size:.2f}")

    print("\tCount of projects by each size:")
    for job_size, count in job_size_counts.items():
        print(f"\t\tJob Size {job_size}: {count}")

    print(f"\tTotal average WSJF number from all projects: {average_wsjf:.2f}")

    print("\tCount of projects by WSJF Range:")
    for wsjf_range, count in wsjf_range_counts.items():
        print(f"\t\tWSJF Range {wsjf_range}: {count}")

def pi_planning_logic_main(file_path):
    # Logic for PI Planning
    filtered_df = filter_excel(file_path)
    analyze_pi_planning(filtered_df, file_path)

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python pi_planning_logic.py <file_path>")
        sys.exit(1)
    file_path = sys.argv[1]
    pi_planning_logic_main(file_path)
