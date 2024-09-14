import pandas as pd
import re

# Define the path to your Excel file (make sure it's in the same directory as the script or provide the full path)
excel_file_path = 'records_embase.xlsx'

# Load the Excel file into a DataFrame
df = pd.read_excel(excel_file_path)

# List of AI-related keywords (including plural and non-plural forms)
ai_keywords = [
    ("artificial intelligence", 1),
    ("artificial intelligences", 1),
    ("artificial-intelligence", 1),
    ("machine learning", 2),
    ("machine learnings", 2),
    ("machine-learning", 2),
    ("automated", 2),
    ("automatic", 2),
    ("LASSO", 2),
    ("Gaussian", 2),
    ("deep learning", 3),
    ("deep-learning", 3),
    ("deep learnings", 3),
    ("data driven", 3),  
    ("big data", 3),      
    ("data-driven", 3),   
    ("learning", 3),
    ("supervised learning", 4),
    ("supervised-learning", 4),
    ("supervised learnings", 4),
    ("unsupervised learning", 5),
    ("unsupervised learnings", 5),
    ("reinforcement learning", 6),
    ("reinforcement learnings", 6),
    ("neural network", 7),
    ("neural-network", 7),
    ("neural networks", 7),
    ("deep neural", 7),
    ("natural language processing", 8),
    ("natural language processings", 8),
    ("computer vision", 9),
    ("computer visions", 9),
    ("data mining", 10),
    ("data-mining", 10),
    ("data minings", 10),
    ("predictive analytics", 11),
    ("predictive analytics", 11),
    ("big data", 12),
    ("big datas", 12),
    ("feature engineering", 13),
    ("feature engineerings", 13),
    ("model training", 14),
    ("model trainings", 14),
    ("algorithm", 15),
    ("algorithms", 15),
    ("pattern", 15),
    ("pattern recognition", 15),
    ("classification", 15),
    ("classified", 15),
    ("robot", 16),
    ("robotic", 16),
    ("robot-assisted", 16),
    ("robotic-assisted", 16),
    ("robotics-assisted", 16),
]

# Dictionary mapping medical condition keywords to labels
medical_condition_mapping = {
    "tremor": 16,
    "tremors": 16,
    "movement disorder": 17,
    "movement disorders": 17,
    "parkinson": 18,
    "parkinsons": 18,
    "parkinsonian": 18,
    "dystonia": 19,
    "dystonias": 19,
    "huntington": 20,
    "huntingtons": 20,
    "essential tremor": 21,
    "essential tremors": 21,
    "ataxia": 22,
    "ataxias": 22,
    "progressive supranuclear palsy": 23,
    "progressive supranuclear palsies": 23,
    "psp": 24,
    "msa": 25,
    "tics": 26,
    "DBS": 27,
    "deep brain stimulation": 27,
}

# Initialize lists to store flags and labels
ai_flags = []
medical_flags = []
ai_labels = []
medical_condition_labels = []

# Iterate through the titles in the "Title" column, starting from row 1 (index 0)
for title in df['Title']:
    if isinstance(title, str):  # Check if the title is a string
        title_lower = title.lower()
        
        # Initialize variables to track flags and labels for the current title
        flag_ai = 2  # Default 'No' flag for AI
        ai_label = None
        flag_medical = 2  # Default 'No' flag for medical condition
        medical_label = None
        
        # Check for AI-related keywords
        for keyword, label in ai_keywords:
            if re.search(r'\b{}\b'.format(re.escape(keyword.lower())), title_lower):
                flag_ai = 1  # 'Yes' flag for AI
                ai_label = label
                break
        
        # Check for medical condition keywords
        for keyword, label in medical_condition_mapping.items():
            if re.search(r'\b{}\b'.format(re.escape(keyword.lower())), title_lower):
                flag_medical = 1  # 'Yes' flag for medical condition
                medical_label = label
                break
        
        ai_flags.append(flag_ai)
        medical_flags.append(flag_medical)
        ai_labels.append(ai_label)
        medical_condition_labels.append(medical_label)
    else:
        # Handle missing or NaN values here (you can skip or set default values)
        ai_flags.append(2)
        medical_flags.append(2)
        ai_labels.append(None)
        medical_condition_labels.append(None)

# Add the flags to new columns named "AI Include" and "Medical Condition Include" starting from row 1 (index 0)
df.insert(loc=df.columns.get_loc("Title"), column='AI Include', value=ai_flags)
df.insert(loc=df.columns.get_loc("Title") + 2, column='AI Label', value=ai_labels)

# Insert the medical condition flags and labels after the "AI Label" column
df.insert(loc=df.columns.get_loc("AI Label") + 2, column='Medical Condition Include', value=medical_flags)
df.insert(loc=df.columns.get_loc("AI Label") + 4, column='Medical Condition Label', value=medical_condition_labels)

# Create the aggregate include column based on the conditions
df['Aggregate Include'] = [1 if ai == 1 and medical == 1 else 2 for ai, medical in zip(ai_flags, medical_flags)]

# Save the updated DataFrame back to the Excel file
df.to_excel(excel_file_path, index=False)

print("Flagging complete. Results saved to the Excel file:", excel_file_path)
