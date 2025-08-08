import pandas as pd
import re
from collections import Counter

def extract_group_counts(file_path, keyword='Groups'):
    df = pd.read_excel(file_path, sheet_name='Input Data sheet', engine='openpyxl')
    comments_series = df['Additional comments'].fillna('') + ' ' + df['Comments and Work notes'].fillna('')
    pattern = rf'{keyword}\s*:\s*\[code\]<I>(.*?)</I>\[/code\]'
    group_counter = Counter()

    for comment in comments_series:
        matches = re.findall(pattern, comment)
        for match in matches:
            groups = [group.strip() for group in match.split(',')]
            group_counter.update(groups)

    return group_counter

if __name__ == "__main__":
    file_path = "coding_challenge_test.xlsx"
    keyword = "Groups"  # You can change this to "Group Names" or any other keyword
    group_counts = extract_group_counts(file_path, keyword)

    with open("group_output.txt", "w") as f:
        for group, count in group_counts.items():
            f.write(f"{group}\t{count}\n")

    print("Group counts written to group_output.txt")

