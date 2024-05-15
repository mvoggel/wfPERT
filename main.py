import pi_planning_logic
import testing
#import keyboard

def get_user_choice():
    print("\nWhat would you like to do? \nType a number and hit enter, or to exit hit enter two times:")
    print("1. Give summary stats")
    print("2. PI Planning logic handling")
    print("3. Run a PERT analysis")

    choice = input("\nEnter your choice (1/2/3): ")
    return choice

def summary_stats(file_path):
    # Logic for summary stats without filtering
    df = testing.read_excel_file(file_path)

    duration_stats, hours_stats, expected_time_days, expected_time_hours = testing.calculate_statistics(df)

    duration_values = df['Actual Duration (days)'].tolist()
    hours_values = df['Actual Hours'].tolist()

    duration_variance, duration_std_dev, hours_variance, hours_std_dev = testing.calculate_variance_and_std_dev(duration_values, hours_values)

    duration_95_confidence_stat = testing.calculate_95_confidence_stat(duration_variance)
    hours_95_confidence_stat = testing.calculate_95_confidence_stat(hours_variance)

    testing.print_results(duration_stats, hours_stats, expected_time_hours, expected_time_days, duration_variance, duration_std_dev, hours_variance, hours_std_dev, duration_95_confidence_stat, hours_95_confidence_stat)

    # Further summary options
    further_summary = input("Would you like a further summary? (yes/no): ")
    if further_summary.lower() == "yes":
        further_summary_choice = input("Select further summary:\n1. By Business Category\n2. By Job Size\nEnter your choice (1/2): ")

        if further_summary_choice == "1":
            business_category_summary(df)
        elif further_summary_choice == "2":
            job_size_summary(df)
        else:
            print("Invalid choice.")

def business_category_summary(df):
    # Summary by Business Category
    business_categories = df['Business Category'].unique()
    for category in business_categories:
        filtered_df = df[df['Business Category'] == category]
        print(f"\nSummary for Business Category: {category}")
        duration_stats, hours_stats, expected_time_days, expected_time_hours = testing.calculate_statistics(filtered_df)

        duration_values = filtered_df['Actual Duration (days)'].tolist()
        hours_values = filtered_df['Actual Hours'].tolist()

        duration_variance, duration_std_dev, hours_variance, hours_std_dev = testing.calculate_variance_and_std_dev(duration_values, hours_values)

        duration_95_confidence_stat = testing.calculate_95_confidence_stat(duration_variance)
        hours_95_confidence_stat = testing.calculate_95_confidence_stat(hours_variance)

        testing.print_results(duration_stats, hours_stats, expected_time_hours, expected_time_days, duration_variance, duration_std_dev, hours_variance, hours_std_dev, duration_95_confidence_stat, hours_95_confidence_stat)

def job_size_summary(df):
    # Summary by Job Size
    job_sizes = sorted(df['Job Size'].unique())
    for size in job_sizes:
        filtered_df = df[df['Job Size'] == size]
        print(f"\nSummary for Job Size: {size}")
        duration_stats, hours_stats, expected_time_days, expected_time_hours = testing.calculate_statistics(filtered_df)

        duration_values = filtered_df['Actual Duration (days)'].tolist()
        hours_values = filtered_df['Actual Hours'].tolist()

        duration_variance, duration_std_dev, hours_variance, hours_std_dev = testing.calculate_variance_and_std_dev(duration_values, hours_values)

        duration_95_confidence_stat = testing.calculate_95_confidence_stat(duration_variance)
        hours_95_confidence_stat = testing.calculate_95_confidence_stat(hours_variance)

        testing.print_results(duration_stats, hours_stats, expected_time_hours, expected_time_days, duration_variance, duration_std_dev, hours_variance, hours_std_dev, duration_95_confidence_stat, hours_95_confidence_stat)


def run_pert_analysis(file_path):
    # Run PERT analysis
    df = testing.read_excel_file(file_path)
    testing.main(df)

def pi_planning_logic_handler(file_path):
    pi_planning_logic.pi_planning_logic_main(file_path)

def main():
    FILE_PATH = "/Users/matthewvoggel/Desktop/Pert Test2.xlsx"  # Path to Excel file
    while True:  # Loop to allow user to perform multiple actions
        choice = get_user_choice()

        if choice == '1':
            summary_stats(FILE_PATH)
        elif choice == '2':
            pi_planning_logic_handler(FILE_PATH)
        elif choice == '3':
            run_pert_analysis(FILE_PATH)
        else:
            print("Invalid choice. Please enter a valid option.")

        print("\nWould you like to return to the main menu? (Type yes, or press Enter to exit)")
        response = input()
        if response.lower() != 'yes':
            break

if __name__ == "__main__":
    main()
