# Global Soil Pollution Analysis
# Soil Pollution Analysis - Excel VBA

## Overview
This project analyzes a synthetic dataset related to soil pollution and associated diseases.

The dataset includes environmental, agricultural, and demographic variables and is designed to explore relationships between soil contamination and human health outcomes.

The project was developed using Microsoft Excel and VBA and focuses on data cleaning, visualization, statistical analysis, forecasting, and interactive data querying.

## Project Objectives:
1. Data Cleaning and Transformation
2. Data Visualization
3. Statistical Analysis
4. Data Input and Querying
5. Highlighting Key Insights and Ecological Risk

## Dataset
This dataset contains 3000 synthetic records simulating real-world scenarios of soil pollution and related diseases. It captures environmental, agricultural, and demographic variables to analyze correlations between soil contamination and human health outcomes. The data was generated using probabilistic models and domain knowledge, making it suitable for exploratory analysis, machine learning, and environmental health research.
Source: https://www.kaggle.com/datasets/khushikyad001/soil-pollution-and-associated-health-impacts/data

The project uses variables related to:
* Countries
* Dates
* Pollutant concentration
* Disease severity
* Environmental conditions
* Demographic information


### 1. Data Cleaning
The VBA macros are used to:
* Remove unnecessary columns
* Remove duplicate records
* Handle missing values
* Sort records by date

### 2. Data Visualization
The project includes visualizations exploring:

* Soil-pollution impacts across different age groups
* Countries and pollution types

These charts provide a visual overview of patterns within the dataset.

### 3. Statistical Analysis
VBA is used to calculate statistical values such as the:

* Mean pollutant concentration

The calculated results are displayed through Excel message boxes.

### 4. Forecasting
The project includes a VBA forecasting function that estimates a future pollutant-concentration value based on the existing dataset.

The forecast result is displayed using a message box.

### 5. Data Querying and Input
Interactive functionality allows users to:

* Find pollutant data for a specific country and date
* Add new records to the dataset

This provides a simple interface for interacting with the data without manually searching through the worksheet.

### 6. Ecological Risk Highlighting
The project identifies pollutant concentrations above the defined threshold of 151.5 mg/kg.

Values above this threshold are highlighted as high risk.

Disease severity is also visually categorized:

| Disease Severity | Highlight |
| ---------------- | --------- |
| Mild             | Green     |
| Moderate         | Yellow    |
| Severe           | Red       |

The VBA macro also calculates the percentage of records exceeding the pollutant-concentration threshold.

## VBA Components
The project contains VBA procedures for:
* Removing unnecessary columns
* Removing duplicates
* Handling missing values
* Sorting dates
* Calculating statistics
* Forecasting future pollutant concentrations
* Highlighting ecological-risk thresholds
* Querying and entering data through user forms

## Tools & Technologies
* Microsoft Excel
* VBA

1. Download the Excel workbook.
2. Open the workbook in Microsoft Excel.
3. Enable macros when prompted.
4. Open the VBA interface if you want to inspect or modify the code.
5. Run the required macro or use the provided UserForm functionality.
6. Review the generated results, charts, and highlighted values.


## Authors
Badis Zammouri
Mohamed Chandoul
