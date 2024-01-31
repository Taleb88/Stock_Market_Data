import pandas as pd

fundamentals_df = pd.read_excel('master.xlsx',
                                sheet_name="fundamentals")
prices_df = pd.read_excel('master.xlsx',
                          sheet_name="prices")
prices_split_adjusted_df = pd.read_excel('master.xlsx',
                                         sheet_name="prices-split-adjusted")
securities_df = pd.read_excel('master.xlsx',
                              sheet_name="securities")

#print(fundamentals_df.head())
#print(fundamentals_df.tail())

# creating fundamentals condensed dataframe
fundamentals_condensed_df = pd.DataFrame()
id = fundamentals_df.iloc[:,0]
fundamentals_condensed_df['ID'] = id.copy()
ticker_symbol = fundamentals_df.iloc[:,1]
fundamentals_condensed_df['Ticker Symbol'] = ticker_symbol.copy()
period_ending = fundamentals_df.iloc[:,2]
fundamentals_condensed_df['Period Ending'] = period_ending.copy()
accounts_payable = fundamentals_df.iloc[:,3]
fundamentals_condensed_df['Accounts Payable'] = accounts_payable.copy()
accounts_receivable = fundamentals_df.iloc[:,4]
fundamentals_condensed_df['Accounts Receivable'] = accounts_receivable.copy()
gross_profit = fundamentals_df.iloc[:,25]
fundamentals_condensed_df['Gross Profit'] = gross_profit.copy()
intangible_assets = fundamentals_df.iloc[:,27]
fundamentals_condensed_df['Intangible Assets'] = intangible_assets.copy()
interest_expense = fundamentals_df.iloc[:,28]
fundamentals_condensed_df['Interest Expense'] = interest_expense.copy()
investments = fundamentals_df.iloc[:,30]
fundamentals_condensed_df['Investments'] = investments.copy()
liabilities = fundamentals_df.iloc[:,31]
fundamentals_condensed_df['Liabilities'] = liabilities.copy()
long_term_debt = fundamentals_df.iloc[:,32]
fundamentals_condensed_df['Long-Term Debt'] = long_term_debt.copy()
long_term_investments = fundamentals_df.iloc[:,33]
fundamentals_condensed_df['Long-Term Investments'] = long_term_investments.copy()
minority_interest = fundamentals_df.iloc[:,34]
fundamentals_condensed_df['Minority Interest'] = minority_interest.copy()
for_year = fundamentals_df.iloc[:,75]
fundamentals_condensed_df['For Year'] = for_year.copy()
earnings_per_share = fundamentals_df.iloc[:,76]
fundamentals_condensed_df['Earnings Per Share'] = earnings_per_share.copy()
estimated_shares_outstanding = fundamentals_df.iloc[:,77]
fundamentals_condensed_df['Estimated Shares Outstanding'] = estimated_shares_outstanding.copy()


correct_year = []

for x in fundamentals_condensed_df['Period Ending']:
    if '2012' in fundamentals_condensed_df['Period Ending']:
        correct_year.append(x)
    elif '2013' in fundamentals_condensed_df['Period Ending']:
        correct_year.append(x)
    elif '2014' in fundamentals_condensed_df['Period Ending']:
        correct_year.append(x)
    elif '2015' in fundamentals_condensed_df['Period Ending']:
        correct_year.append(x)

fundamentals_condensed_df['For Year'] = correct_year

'''
# if '2012', '2013', '2014', '2015', '2016' in Period Ending
def earnings_per_share_blank_value(df):
    blank = []
    for x in df['Period Ending']:
        if '2014' in df['Period Ending']:
            blank.append(x)
'''

# remove timestamp from period ending values
fundamentals_condensed_df['Period Ending'] = \
    fundamentals_condensed_df['Period Ending'].astype(str).str[:11] #convert to string format and remove timestamp

# created new workbook containing fundamentals condensed dataframe
fundamentals_condensed_df.to_excel('fundamentals_condensed_df.xlsx', index=False)