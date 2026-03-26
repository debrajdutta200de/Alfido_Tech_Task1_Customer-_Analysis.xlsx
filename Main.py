import pandas as pd

# Creating a sample dataset based on the task requirements
data = {
    'CustomerID': range(101, 121),
    'Gender': ['Male', 'Female', 'Female', 'Male', 'Female', 'Male', 'Male', 'Female', 'Male', 'Female'] * 2,
    'Age': [23, 45, 31, 22, 54, 38, 41, 29, 33, 47, 25, 50, 32, 21, 55, 39, 42, 30, 34, 48],
    'Annual_Income_k$': [15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 60, 62, 63, 64, 65, 70, 71, 72, 75, 80],
    'Spending_Score': [39, 81, 6, 77, 40, 76, 6, 94, 3, 72, 49, 15, 41, 99, 11, 75, 44, 1, 5, 14],
    'Purchase_Frequency': [5, 12, 2, 15, 4, 11, 2, 18, 1, 10, 8, 3, 6, 20, 2, 14, 7, 1, 2, 3]
}

df = pd.DataFrame(data)

# Customer Segmentation Logic
def segment_customer(row):
    if row['Spending_Score'] > 70 and row['Purchase_Frequency'] > 10:
        return 'High Value (Loyal)'
    elif row['Spending_Score'] < 30:
        return 'At Risk (Low Engagement)'
    else:
        return 'Average/Potential'

df['Customer_Segment'] = df.apply(segment_customer, axis=1)

# Summary for Recommendation Sheet
recommendations = {
    'Strategy': [
        'Loyalty Program',
        'Re-engagement Campaign',
        'Personalized Offers',
        'Referral Discounts',
        'Upselling Strategy'
    ],
    'Actionable Steps': [
        'Give 10% extra discount to High Value customers to maintain loyalty.',
        'Send "We Miss You" emails with coupons to At Risk customers.',
        'Suggest products based on previous purchase categories.',
        'Encourage loyal customers to refer friends for a reward.',
        'Show premium products to Average spenders to increase their spending score.'
    ]
}
df_rec = pd.DataFrame(recommendations)

# Saving to Excel with multiple sheets
with pd.ExcelWriter('Alfido_Tech_Task1_Customer_Analysis.xlsx') as writer:
    df.to_excel(writer, sheet_name='Analysis_Data', index=False)
    df_rec.to_excel(writer, sheet_name='Recommendations', index=False)
