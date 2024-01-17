# Efrain Cortes MIS-15 01

import numpy as np # useful for many scientific computing in Python
import pandas as pd # primary data structure library
import matplotlib as mpl
import matplotlib.pyplot as plt 

# I didn't know which way you wanted the file in df_can, you can get the file to run by replacing with the excel file on your computer in pd.read_excel()
df_can = pd.read_excel("c:/Users/Efrain/Downloads/Generic University Data Updated.xlsx")

# Variable names for each excel spreadsheet
sheetOne = pd.read_excel("c:/Users/Efrain/Downloads/Generic University Data Updated.xlsx",'Student Data') 
sheetTwo = pd.read_excel("c:/Users/Efrain/Downloads/Generic University Data Updated.xlsx",'Schools')       
sheetThree = pd.read_excel("c:/Users/Efrain/Downloads/Generic University Data Updated.xlsx",'Programs')    


# Open Student Data Sheet
# Print the dimension of the dataset
rows_count = sheetOne.shape[0]
col_count = sheetOne.shape[1]
print('Rows:', rows_count, '\nColumns:', col_count)

# Print the top 5 rows of the dataset and the bottom 5 rows
print("\nPrinting Top 5 Rows:\n" , sheetOne.head())
print("\nPrinting Bottom 5 Rows:\n" , sheetOne.tail())


# Drop [Aid type, ID, Pell Grant Amount, Postal Code, City} from the sheet
print("Printing Top 5 Rows After Info Dropped:\n" , sheetOne.drop(['Aid Type', 'ID', 'Pell Grant Amount', 'Postal Code', 'City'], axis=1))

# Display rows with 'Scholarship' for students in California.
# Include Academic Year, GPA, Ethnicity, Gender, Name, Program, SAT, Scholarship Amount, School, and State in the output.
print("\n\n\n\nPrinting based on conditions: Aid Type - Scholarship, State - California")
print(sheetOne[(sheetOne['Aid Type'] == 'Scholarship') & (sheetOne['State'] == 'California')].drop(['Aid Type' , 'City' , 'ID' , 'Pell Grant Amount' , 'Postal Code'], axis=1))


# Display only rows pertaining to female students with 'Scholarship.'
# Include Academic Year, GPA, Ethnicity, Gender, Name, Program, SAT, Scholarship Amount, School, and State in the output.
print("\n\n\n\nPrinting based on conditions: Aid Type - Scholarship, Gender - Female")
print(sheetOne[(sheetOne['Aid Type'] == 'Scholarship') & (sheetOne['Gender'] == 'female')].drop(['Aid Type' , 'City' , 'ID' , 'Pell Grant Amount' , 'Postal Code'], axis=1))

# Schools Sheet
# Create 'Total' Column 
sheetTwo['Total'] = sheetTwo.drop(['School'], axis=1).sum(axis=1)

# Define variable, year, which contains list of all the years
year = list(map(str, range(2008,2019)))

# Plot a line graph for the School of Business and Management
sheetTwo.set_index('School',inplace=True)
businessScholarship = sheetTwo.loc['Business and Management',year]
businessScholarship.plot(kind='line')
plt.title('School of Business Scholarship Funding')
plt.xlabel('Years')
plt.ylabel('Scholarship Funding (in millions)')
plt.yticks([300000,600000,900000,1200000,1500000,1800000])
plt.show()

# Plot a line graph of scholarship amounts from the five schools, which one has the highest amount of scholarship in 2018
schoolSheet = sheetTwo.loc['Health Sciences':'Engineering', year].transpose()
schoolSheet.plot(kind='line')
plt.title('All Scholarships')
plt.xlabel('Years')
plt.ylabel('Scholarship Funding (in millions)')
plt.text(4, 1554781, 'Highest Funding in 2018')
plt.show()


# Program Sheet
# Create 'Total' Column 
sheetThree['Total'] = sheetThree.drop(['Program'], axis=1).sum(axis=1)

# Arrange the DataFrame by the designated column in descending order.
sheetThreeSorted = sheetThree.sort_values(by="Total", ascending=False)

# Generate a new dataframe containing only the first 5 data rows.
sheetThreeTopFive = sheetThreeSorted.head()
sheetThreeTopFive.set_index('Program',inplace=True)
sheetThreeTopFive = sheetThreeTopFive.loc['Applied Math':'Comparative History of Ideas',year].transpose()

# Plot the stacked area graph for the Top 5 program sheet data
sheetThreeTopFive.plot(kind='area', stacked=True)
plt.title('Top 5')
plt.xlabel('Years')
plt.ylabel('Scholarship Funding (in millions)')
plt.show()


# Generate a new dataframe for sorting the designated columns: 'Computer Science,' 'Information Management,' and 'Business Administration.'
sheetThreeSpecificSort = sheetThree
sheetThreeSpecificSort.set_index('Program',inplace=True)
sheetThreeSpecificSort = sheetThreeSpecificSort.loc[['Computer Science','Information Management','Business Administration'],year].transpose()


# Plot the histogram 
sheetThreeSpecificSort.plot(kind='hist', stacked=True)
plt.title('Sort: CS IM BA')
plt.xlabel('Years')
plt.ylabel('Scholarship Funding (in millions)')
plt.show()