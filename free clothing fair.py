import pandas as pd
import matplotlib.pyplot as plt

# Step 1: Load the Dataset
# Replace 'Free Clothing Fair Data 2024.xlsx' with the path to your file
before = pd.read_excel("Free Clothing Fair Data 2024.xlsx", sheet_name="before")
after = pd.read_excel("Free Clothing Fair Data 2024.xlsx", sheet_name="after fair store")

# Step 2: Clean the Data
# Fill missing values with 0 (or handle as needed)
before.fillna(0, inplace=True)
after.fillna(0, inplace=True)

# Step 3: Group by 'Item' and Calculate Counts and Weights

# Group by 'Item' to calculate count (occurrences) and sum of weights
# Group by 'Item' and calculate sum of 'Weight'
before_by_item = before.groupby('Item').agg({'Weight': 'sum'}).reset_index()

# Calculate the count of occurrences (how many times each item appears)
before_by_item['Count'] = before.groupby('Item').size().reset_index(name='Count')['Count']


# Step 4: Calculate Key Metrics
# Total items before and after the event (based on the count of items)
before_total_items = 937.9
after_total_items = 136.7
donated_items = 801.2

print(f"Total Items Before Event: {before_total_items}")
print(f"Total Items After Event: {after_total_items}")
print(f"Total Items Donated: {donated_items}")

# Step 5: Visualize the Data
# Bar chart for items before and after by item type (based on count)
plt.figure(figsize=(12, 6))
plt.bar(before_by_item['Item'], before_by_item['Weight'], alpha=0.7, label="Before Event")
plt.xlabel("Item Type")
plt.ylabel("Weight")
plt.title("Items Before Event by Item Type (Weight)")
plt.xticks(rotation=45)
plt.legend()
plt.tight_layout()
plt.show()


plt.figure(figsize=(6, 6))
labels = ['Donated', 'Remaining']
sizes = [donated_items, after_total_items]
plt.pie(sizes, labels=labels, autopct='%1.1f%%', colors=['#ff9999', '#66b3ff'])
plt.title("Donated vs. Remaining Items")
plt.show()

# Step 6: Group by Size and Plot Size Distribution
before_by_size = before.groupby('Size').size().reset_index(name='Count')

plt.figure(figsize=(12, 6))
plt.bar(before_by_size['Size'], before_by_size['Count'], alpha=0.7, label="Before Event")
plt.xlabel("Size")
plt.ylabel("Boxes Count")
plt.title("Size Distribution Before the Event")
plt.legend()
plt.tight_layout()
plt.show()

# Step 7: Save the Results (Optional)
# Save summary results to a new Excel file
summary = pd.DataFrame({
    'Item': before_by_item['Item'],
    'Before Count': before_by_item['Count'],
})
summary.to_excel("Free_Clothing_Fair_Summary.xlsx", index=False)
print("Summary saved to 'Free_Clothing_Fair_Summary.xlsx'")
