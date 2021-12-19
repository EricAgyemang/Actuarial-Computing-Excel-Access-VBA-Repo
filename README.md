# Actuarial-Computing-Excel-Access-VBA-Repo
  Actuarial Computing (MAT 253 , ISu)
  
# Using VLOOKUPS
This code lab focuses on using VLOOKUPS to fill out tables listed under for number of claims, referencing the first table on the data tab. It involves making use of values in rows in order to complete the 3rd parameter of the VLOOKUP function.  It also involvs the correct use of absolute and relative cell referencing so that can be copied the same function across the entire YELLOW area.

# Using HLOOKUP
This code lab also focuses on using  HLOOKUP to fill out tables using data from a second table of the Data 1 tab.  For the 3rd parameter of the HLOOKUP, we make use of the MATCH function with the appropriate match key and array reference to the vector with a list of available years

# LOSS AND PREMIUM CALCULTIONS
The code involve following the instructions below:
   - Loss Amount and Pure Premium columns should be formatted with US currency and two decimal places.
   - Claim Frequency columns should be formated in percetages with 2 decimal places.
   - Exposure Units should be formated with a comma and no decimal places.
   - create a conditional format for column G if the growth rate less than  0%.  
   - Cells that satisfy this condition format should have their font in BOLD and the pattern should fill with red.  
   - Add a second conditional format if the growth rate is  excees 5%.  
   - Cells that satisfy this condition should have their font in BOLD and their pattern should fill with green.
   - Create a data filter on the data on the Problem 2 tab.  
   - You should select rows where STAT column is "MS" and where the COV is "COLL". 
   - Please keep the filter on when you save your file to return to me.

# GRAPHS
In this code lab, we Create graphs which show the actual claim frequency and the actual claim severity on the y axis.  
Because the scale for each of these series is so different, we use two different axes to show the different series.  
The x axis show the period # (cloumn A).  Each series display as points, with connecting lines.  
Each series is labled as frequency or severity as appropriate.  

# USING IF STATEMENTS
Using IF statements, we calculate the Actuarial present value for each of the people in the list on the "problem 1" tab.
    - The APV formula = Face Value * Ax
    - The Ax varies by sex and smoker status and can be found on the 4 tabs for each case.
In order to check the answer, the result of the first policy should have APV = 1,1238.0
On the “Problem 1” tab, column A contains a text string that’s a concatenation of 4 different fields: Policy_Num, 
Effective_Date, Expiration_Date, Premium.  Use comma (,) as delimiter to separate them into 4 columns. You can use any tool or f
unctions within Excel to do it. 

# PIVOT TABLE
WE set up a PivotTable report in a new worksheet call “Problem 1” from the data on the 'Collection' tab (range A1:D2771).  Put the 'Number
of Collection' in the row labels, and create 4 columns:
1.      Sum of Premium
2.      Sum of Loss
3.      Loss Ratio = Loss / Premium
4.      Policy count, display as % of column.


# REGRESSION
On the “Regression” tab, use simple linear regression (y=a+bx) technique to predict a person's weight using their height.  You can use any methods 
that are available in Excel to obtain the parameter estimates.

# CREATING QUERY
  - Create a Query in access that returns a list of all the policy numbers in the states of IL, IN, OH, and MI.  Save this query with name “Problem 1a”.  
    In the criteria, use in (“IL”, “IN”, “OH”, “MI”)
  - Merge “PolicyTable” and “LossTable” by POLICY_NUM (you will need to use an outer join, retaining all policy from PolicyTable). 
  - Create a cross tab query based on the “PolicyTable” table with the following properties:
           * Row Heading:  Class
           *  Column Heading:  Deductible
           *  Value:  Average Rate
   The results should only reflect the following four states:  IL, MI, OH, WI.  Hint: the “total” selection should be selecting “where” for variable State. 
   And you can use the in () statement.



# COMPREHENSIVE PROJECT
  
  - Project Information
This project is intended to cover many of the topics that you’ve worked on during the course of this project.  The project is due prior to the final 
class on April 24, 2018.  
The project involves work in Excel, and Access.  In general, the processing relies on processing in Access, then work in Excel.  
The intent with this project is to provide less specific direction on the project compared to the homework.  Most, if not all, the formulas used in this example 
can be found in the examples and homework that we’ve covered so far in class.


 # REAL WORLD SCENARIO
- You are a pricing actuary for ABC Insurance Company, a small personal lines auto insurer with premium revenue of approximately $300M annually.
  One of your job responsibilities is to develop periodic rate level indications, as well as adjustments to your rating factors.  Your boss has asked you to put 
  together a process to streamline the indications process for developing indicated rates for 2011.  To do so, he has provided the following instructions as well as
  a shell of what he would like the spreadsheet to look like.

- He has also asked you to provide a separate way for him to keep an eye on pure premium trends in all states, and compare to countrywide (CW) trends.  He’d like 
  a simple point and   
- click method to do this, so you’ve suggested a PivotChart for this purpose.

- Developing a rate indication at ABC involves a few steps including:
  •	Trend Analysis
  •	Development of Loss Projection Factors based on Trends
  •	Developing indicated Deductible and Class (Age & Sex) Factors
  •	Development of Investment Yield
  •	Development of overall rate indication

- To develop the rate indications, you have been given the following information:
  •	The IT department has provided detailed premium and loss information for all policies 2007-2009 in a fixed width text file.  This file has about 1 million 
    records, so it must be first processed in Access.
  •	You also have a copy of the latest Fast Track industry trend data in an Access database
  •	You have an Excel spreadsheet with the company’s stock holdings and purchases, as well as historical prices for those stocks over the past 4 years.
  
-Your rate indications process will include the following output (explained in more detail below):
  •	An Access database that has queries that output data that can be copied into Excel for each state. 
  •	An Excel spreadsheet that shows the calculation of the average investment yield for 2007-2009.
  •	An Excel spreadsheet that calculates the indicated rate change, after pasting the output of the access queries and investment yield into it.  
    This spreadsheet should allow the user to paste the access output for another state into Excel, and automatically generate the indicated rate without any 
    additional updates.
  •	An Excel spreadsheet with a PivotChart that displays both the CW trend and the State trend.
    There is an example of what the output from the rate indications worksheet should look like.


# MICROSOFT ACCESS PART
An Access database is provided.  That database already contains a table named TrendData, which has the industry trend data.
You have also been provided with detailed policy data on policydata.txt.  The layout for the text file is below:
    Pos	Field
    1-2	Keys
    3-4	State
    5-8	Deductible
    9-14	Class Code
    15-18	Year
    19-24	Premium
    25	Indicator whether policy had claim
    26-35	Claim Amount

** Note on Keys field**

     Please use Access to add a primary key.  The keys field in the input dataset is truncated.  (Thus not unique to each record.)  However, it will not impact your calculations.
     You should import the text file with the policy data into an Access table.

In Access, you should create queries that outputs the following information:

- Company Premium/Loss Information:
  STATE  (Group By)
  YEAR  (Group By)
  DEDUCT  (Group By)
  CLASS  (Group By)
  Policy Count   (Count)
  PREM (Sum)
  CLAIM_IND (Sum)
  LOSS_AMOUNT (Sum)
  
- You should set the query to have a where clause for the state.  You can change the state to whichever state you are working on.
  Industry Fast Track Trend Information:
  STATE (Group by)
  YYYYQ (Group by)
  Cov (Group by)
  CW_CARYEARS (Sum)
  CW_PDCOUNT  (Sum)
  CW_PDAMT  (Sum)
  STATE_CARYEARS  (Sum)
  STATE_PDCOUNT  (Sum)
  STATE_PDAMT  (Sum)

The CW fields are summaries based on all data for all states.  The STATE summary fields are sums of the fields for the particular state.  Again, you should set up the query for the Where clause to specify the state that is to be outputted.

Note that to get both CW summaries and STATE summaries on the same query, you’ll have to merge the output of two separate queries (one at the state level, and one at the CW level) and merge the results by YYYQ and COV.


# Excel Investment Yield Worksheet

The spreadsheet provided has two tables.  One table has the stock prices over time for stocks on the S&P 500.  ABC Company owns a subset of those stocks.  The investment department has a provided a summary of the stocks held at the beginning of year (BOY) 2006, as well as stocks purchased on 1/1/2007, 1/1/2008, and 1/1/2009.
You need to calculate the investment yield for 2007, 2008, and 2009, and the arithmetic average of the 3 year yield.   A demonstration of the calculation is included the handout.  You should fill out the spreadsheet on the Investment Yield Calculation of the worksheet.
The value that you calculate in this worksheet will be entered in the Rate Indications worksheet.


# Excel Rate Indications Worksheet
The output from Access should be pasted into the Input Data tab of the worksheet.  Feel free to add any index columns to this tab that might be useful to you later on.  You should also be able to input the State Name on that tab and have the resulting State name flow to all the Worksheet headers in the worksheet (so if you paste data for a new state, you only have to change the state name once in the worksheet, rather than having to update every sheet).
Keep in mind that no other changes should be necessary when updating a state.  Think about the possibility of queries for different states returning a different number of rows.  You may need to use larger references to the InputData tables than you would for the state data that is in there already.
Included in the handout is an example of what the excel output should look like for the other worksheet tabs.  I’ve listed some tips on completing each sheet on handout.

# Trends Worksheets
- Get the trend information from the output of the Fast Track query.  Your company uses only industry data for trend analysis, and weights the state experience with the CW experience to develop its trends.
- Use the LINEST and the INTERCEPT formulas to calculate the appropriate values.  Feel free to put index (1,2,3,…) in column A for your X-values. Your Y-values should be Pure Premium column.  Remember, Pure Premium = Loss Amount / Car Years.  Use these values to calculate the fitted values columns.  The annual change is 4 x the slope (for four periods).  Express this as a % trend by dividing the annual amount by the most recent fitted value
- Create a graph as shown in the handout with 4 series, State and CW, fitted and actual.

- Create a trend exhibit for all coverages shown.  Keep in mind that you can copy of the first tab you complete by right clicking on it, and say move or copy, then make a copy.  If   
  you code the first tab right, you should just be able to copy it, change the coverage reference, and you won’t have to repeat any of the remaining work.

- Loss Projection Factor Worksheet
  The trends calculated for each coverage should pull through to this worksheet.  There is a credibility weighting calculation on this spreadsheet.  The credibility given to the  
  experience of a given state is based on the number of claims for that state in the most recent period. (For example, if the state 2010 Q1’s claim count for BI is 123,245; the  
  credibility weight assigned should be 0.4.) Those should be pulled from either the trend worksheets, or the raw data on the input data tab. 
  
  **The formula for weighted trend  = State Trend * Credibility weight +  CW Trend * (1-Credibility weight).**

-You also should include the loss amount for the most recent period.  This is used to calculate a weighted average trend for all coverages (cell H13), based on the   
 state’s coverage distribution.

- At the bottom of the worksheet, the LPF is calculated.  The LPF Formula = 1 + # Years Projected * Selected Trend.  For the selected trend, use the credibility   
  weighted all coverage trend from above.

# Deductible and Age/Sex Factor Worksheets
Get the policy count, premium, and loss information for all three years from the company experience data on the input data tab.  Calculate the loss ratio, indicated change, and indicated rate factors.  The indicated change calculation is shown on the spreadsheet.  The indicator factor = Current Factor x (1 + indicated Change).
One both worksheets, add a conditional format to the indicated change column to highlight cells that have a greater than 10% increase, or less than -10% decrease.

# Indicated Change Worksheet

- Pull the premium and loss information from the company experience data on the input data tab.  Pull the LPF from the Loss Projection Factor tab.  Calculate the projected losses = Actual losses x LPF.  

- Use the Projected loss ratio for the 3 year period in the indicated change formula at the bottom of the worksheet.  Manually input the investment yield from your investment yield worksheet.  For the other values in the formula, use the values in the attached example.


# Excel PivotChart Worksheet

-Your boss would also like a way to keep track of trends, without having to do all the work involved with setting up an indications worksheet.  You’ve agreed to create a PivotChart that shows pure premium trends.

-To generate the source data for this PivotChart, you should be able to use the same query as you used to generate the trend data that you pasted into the Indications worksheet.  The main difference is that you should remove the specific state when executing that query.  The query should return the values for all states, as well as columns that contain the CW values.  Paste the output of the query into a new Excel workbook.

# Creating a PivotChart.  
The PivotChart should have Page fields of Coverage and State.  The time period (YYYQ) should be displayed across the bottom of the chart.  The data elements in the chart area should include the state pure premium, and the CW pure premium.




















