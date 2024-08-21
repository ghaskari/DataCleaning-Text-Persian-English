import pandas as pd

# Step 1: Load the CSV file
df = pd.read_csv('DataSource/cleaned_OpenSubtitles-fa.csv')
print(df)
# Step 2: Split the DataFrame into three parts
# Option A: Even split
part_size = len(df) // 3
part_1 = df.iloc[:part_size]
part_2 = df.iloc[part_size:2*part_size]
part_3 = df.iloc[2*part_size:]

# Option B: Custom index-based split
# part_1 = df.iloc[:index1]
# part_2 = df.iloc[index1:index2]
# part_3 = df.iloc[index2:]

# Option C: Condition-based split
# part_1 = df[df['column_name'] < value1]
# part_2 = df[(df['column_name'] >= value1) & (df['column_name'] < value2)]
# part_3 = df[df['column_name'] >= value2]

# Step 3: Handle each part (Example: save each part to a new CSV file)
part_1.to_csv('part1.csv', index=False)
print(part_1)
part_2.to_csv('part2.csv', index=False)
part_3.to_csv('part3.csv', index=False)
