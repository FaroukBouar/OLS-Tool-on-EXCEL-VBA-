# Excel VBA OLS Regression Tool

This project is an Excel **macro-enabled workbook** that performs **Ordinary Least Squares (OLS) regression** with multiple independent variables.

## Features
- UserForm for selecting the dependent variable (Y) and independent variables (X)  
- Calculates **Beta coefficients**, **R²**, **Adjusted R²**, **t-statistics**, **p-values**, **MSE**, **SSR**, **SSE**, and **SST**  
- Automatically adds a **constant (intercept) term**  
- Supports multiple independent variables (multi-column selection)

## How to Use
1. Open the `.xlsm` workbook.  
2. Go to the **Data** sheet and click **Perform a Linear Regression**.  
3. Click **Select Y** and choose your dependent variable range. **Important:** the first row must contain the variable name.  
4. Click **Select X** and choose the range(s) for your independent variable(s). **Important:** the first row must contain the variable names.  
5. Click **Perform OLS** to run the regression and display the results.
