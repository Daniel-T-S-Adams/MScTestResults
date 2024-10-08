import pandas as pd
import matplotlib.pyplot as plt

# List of Excel sheet file paths 
excel_sheets = ["C:\\Users\\44755\\Documents\\Python Scripts\\MScAptitudeTestAnalysis\\MSc Data Science and AI Math and Stats Quiz Two.xlsx", 
                "C:\\Users\\44755\\Documents\\Python Scripts\\MScAptitudeTestAnalysis\\MSc Data Science and AI Math and Stats Quiz One.xlsx"]

# Define the categories and their respective search patterns
categories = {
    'Calc': 'Points - (Calc)',
    'StatInterp': 'Points - (StatInterp)',
    'Prob': 'Points - (Prob)',
    'FndofMath': 'Points - (FndofMath)',
    'LinAlg': 'Points - (LinAlg)',
    'CodeLogic': 'Points - (CodeLogic)'
}

# Mapping for display labels in the plot
display_labels = {
    'Calc': 'Calc',
    'StatInterp': 'Stat Interp',
    'Prob': 'Prob',
    'FndofMath': 'Foundational Math',
    'LinAlg': 'Lin Alg',
    'CodeLogic': 'Logic & Code'
}

# Function to calculate sums and averages for a given course filter
def calculate_averages_for_course(df, course_name):
    # Find the column that starts with "Let us know"
    course_column = [col for col in df.columns if col.startswith('Let us know')]
    
    if course_column:
        # Filter the dataframe for the specific course
        course_df = df[df[course_column[0]] == course_name]
    else:
        raise KeyError("Could not find a column starting with 'Let us know'")
    
    # Initialize sums and counts for the course
    sums = {category: 0 for category in categories}
    counts = {category: 0 for category in categories}

    # Process each category
    for category, pattern in categories.items():
        # Filter columns that contain the pattern in their name
        filtered_df = course_df.filter(like=pattern)
        
        # Sum all the values in these columns
        sums[category] += filtered_df.sum().sum()  # sum across both axis (columns and rows)
        
        # Count the total number of non-NA values in these columns
        counts[category] += filtered_df.count().sum()

    # Calculate the averages for each category
    labels = []
    averages = []
    for category in categories:
        if counts[category] > 0:
            average = sums[category] / counts[category]
            average = average * 100  # get as a percentage
            labels.append(display_labels[category])  # Add the label
            averages.append(average)  # Add the average value
        else:
            print(f"No data found for '{categories[category]}' in {course_name}.")

    return labels, averages

# Read the Excel sheets and concatenate the data
df_list = [pd.read_excel(sheet) for sheet in excel_sheets]
combined_df = pd.concat(df_list, ignore_index=True)

# Filter and calculate averages for both courses
courses = ['Data Science', 'Applied AI']

for course in courses:
    labels, averages = calculate_averages_for_course(combined_df, course)

    # Plotting the bar chart for each course
    plt.figure(figsize=(10, 6))
    plt.bar(labels, averages, color='blue')
    plt.xlabel('Subjects')
    plt.ylabel('Averages (%)')
    plt.title(f'Averages of Different Subjects for {course}')
    plt.ylim(0, 100)  # Set y-axis limit to 100%

    # Optional: Add gridlines to make it easier to read
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    # Save the plot as an image
    plt.savefig(f'{course}_Averages.png')
