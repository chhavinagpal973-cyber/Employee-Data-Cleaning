# importing necessary libraries

import pandas as pd
import numpy as np

# loading the dataset

df = pd.read_csv("C:\\Users\\enagchh\\Downloads\\employee_dirty_data.csv")
print(df.head())

# checking the missing values

print("Missing values in each column")
print(df.isnull().sum())

# How to handle missing value and fillna

df['salary'] = df['salary'].fillna(df['salary'].mean())

df ['performance_rating'] = df['performance_rating'].fillna(df['performance_rating'].median())

# how to deal with infinite values in data

df.replace([np.inf, -np.inf], np.nan, inplace =True)


# remove duplicates

df.drop_duplicates(inplace=True)

# replace negative salaries

df['salary'] = np.where(df['salary'] < 0, df['salary'].mean(), df['salary'])

salary_mean = df['salary'].mean()
salary_std = df['salary'].std()
lower_bound = salary_mean - (3 * salary_std)
upper_bound = salary_mean + (3 * salary_std)

# remove rows ehere salary is too high or too low
df = df[(df['salary'] >= lower_bound) & (df['salary'] <= upper_bound)]

df.to_csv('C:\\Users\\enagchh\\Downloads\\employee_dirty_data.csv', index=False)

print('Data cleaning completed! Saved as "cleaned_indian_employee_Data.csv" ' )