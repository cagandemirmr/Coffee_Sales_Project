# Coffee Sales Project

# Project Overview

In this project, Excel was chosen as the primary tool due to the dataset's manageable size (under 250 MB) and the ease with which data can be quickly reviewed and edited.

## About the Dataset

The dataset consists of three worksheets: `Orders`, `Customers`, and `Products`.

1. **Orders Worksheet**:
   - Columns: `Customer ID`, `Order Date`, `Product ID`, `Quantity`.
   - Variables:
     - `Customer ID` and `Product ID`: Categorical variables.
     - `Order Date` and `Customer ID`: Cardinal variables due to their diverse data types.
     - `Quantity`: Numerical variable representing the amount of coffee packages.
   - Last row: 1001.

2. **Customers Worksheet**:
   - Columns: `Customer ID`, `Customer Name`, `Email`, `Phone Number`, `Address Line 1`, `City`, `Country`, `Postcode`, `Loyalty Card`.
   - Variables:
     - All columns are ordinal categorical variables, except for `Loyalty Card`, which is a binary categorical variable.
   - Last row: 1001.

3. **Products Worksheet**:
   - Columns: `Product ID`, `Coffee Type`, `Roast Type`, `Size`, `Unit Price`, `Price per 100g`, `Profit`.
   - Variables:
     - `Coffee Type`: Categorical variable with four categories: Ara, Exc, Lib, and Rob coffee beans.
     - `Roast Type`: Categorical variable with three categories: Dark, Light, and Medium roast.
     - `Size`: Represents the package size of products.
     - `Price per 100g` and `Profit`: Numerical variables.
   - Last row: 49.

## Client's Requirements

The client requested the following analyses:

- **Sales Analysis**: Evaluate coffee bean sales over time and assess the impact of package size, roast type, and loyalty card usage.
- **Top Performing Countries**: Identify the top three countries with the highest sales.
- **Top Customers**: Determine the top five customers based on sales.

## Project Process

### Cell Operations

The first step was to consolidate data from all worksheets into a single table to facilitate the creation of a pivot table. Ten new columns were added: `Customer Name`, `Email`, `Country`, `Coffee Type`, `Roast Type`, `Size`, `Unit Price`, `Sales`, `Coffee Name`, and `Loyalty Card`.

- **Adding Customer Information**:
  - The `Customer Name`, `Email`, and `Country` columns were populated using the `XLOOKUP` function and the `Country` worksheet.

    ![image](https://github.com/user-attachments/assets/b90e06f7-d008-4195-8292-ba805b18c287)

- **Adding Product Information**:
  - The `Roast Type`, `Size`, and `Unit Price` columns were filled using the `INDEX` and `MATCH` functions, referencing the `Products` worksheet.

    ![image](https://github.com/user-attachments/assets/ce580bc3-2ff0-4a09-b4ea-51ec5dee9012)

## Sales Calculation

The `Sales` column was created by multiplying the `Size` and `Unit Price` columns to calculate the total sales value.

![image](https://github.com/user-attachments/assets/ed75f034-ba68-402f-8a45-58c8df4bf40b)

To provide a more descriptive name for the coffee types, I used the `IFS` conditional function to generate the long versions of coffee names.

## Creating Pivot Tables

![image](https://github.com/user-attachments/assets/350c9c01-0a2b-4b07-b803-383fb0b73277)

I selected the `Order` table and created a pivot table to display the top sales over time by coffee types, with the timeline grouped by month and year.

![image](https://github.com/user-attachments/assets/1259250c-2852-445b-96a2-b5414e02d378)

A line graph was generated from the pivot table, with customized colors and labels added for better clarity.

![image](https://github.com/user-attachments/assets/6cb2aaca-94e4-48c2-8745-5a7e641e8699)

Additionally, I created another pivot table to summarize the total sales by country, specifically filtering for the top 3 countries with the highest sales.

![image](https://github.com/user-attachments/assets/14f1d18a-4eb5-4c13-a6e2-5ffab85d5515)

Lastly, a pivot table was developed to identify the top 5 customers based on sales.

## Dashboard

Finally, I created a dashboard incorporating the graphs and pivot tables developed earlier, providing a comprehensive overview of the sales data.



