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
fundamentals_condensed_df['ID'] = \
    id.copy()
ticker_symbol = fundamentals_df.iloc[:,1]
fundamentals_condensed_df['Ticker Symbol'] = \
    ticker_symbol.copy()
period_ending = fundamentals_df.iloc[:,2]
fundamentals_condensed_df['Period Ending'] = \
    period_ending.copy()
accounts_payable = fundamentals_df.iloc[:,3]
fundamentals_condensed_df['Accounts Payable'] = \
    accounts_payable.copy()
accounts_receivable = fundamentals_df.iloc[:,4]
fundamentals_condensed_df['Accounts Receivable'] = \
    accounts_receivable.copy()
gross_profit = fundamentals_df.iloc[:,25]
fundamentals_condensed_df['Gross Profit'] = \
    gross_profit.copy()
intangible_assets = fundamentals_df.iloc[:,27]
fundamentals_condensed_df['Intangible Assets'] = \
    intangible_assets.copy()
interest_expense = fundamentals_df.iloc[:,28]
fundamentals_condensed_df['Interest Expense'] = \
    interest_expense.copy()
investments = fundamentals_df.iloc[:,30]
fundamentals_condensed_df['Investments'] = \
    investments.copy()
liabilities = fundamentals_df.iloc[:,31]
fundamentals_condensed_df['Liabilities'] = \
    liabilities.copy()
long_term_debt = fundamentals_df.iloc[:,32]
fundamentals_condensed_df['Long-Term Debt'] = \
    long_term_debt.copy()
long_term_investments = fundamentals_df.iloc[:,33]
fundamentals_condensed_df['Long-Term Investments'] = \
    long_term_investments.copy()
minority_interest = fundamentals_df.iloc[:,34]
fundamentals_condensed_df['Minority Interest'] = \
    minority_interest.copy()
for_year = fundamentals_df.iloc[:,2] # period ending and will slice the month and day as we only want the year value
fundamentals_condensed_df['For Year'] = \
    for_year.copy()
earnings_per_share = fundamentals_df.iloc[:,77]
fundamentals_condensed_df['Earnings Per Share'] = \
    earnings_per_share.copy()
estimated_shares_outstanding = fundamentals_df.iloc[:,78]
fundamentals_condensed_df['Estimated Shares Outstanding'] = \
    estimated_shares_outstanding.copy()


# remove timestamp from period ending values
fundamentals_condensed_df['Period Ending'] = \
    fundamentals_condensed_df['Period Ending'].astype(str).str[:11] #convert to string format and remove timestamp

#remove month and year from for year values
fundamentals_condensed_df['For Year'] = \
    fundamentals_condensed_df['Period Ending'].astype(str).str[:4] #convert to string and grab the year value only

# created new workbook containing fundamentals condensed dataframe
fundamentals_condensed_df.to_excel('fundamentals_condensed_df.xlsx', index=False)


# creating earnings dataframe from fundamentals condensed dataframe
earnings_df = pd.DataFrame()
id = fundamentals_condensed_df.iloc[:,0]
earnings_df['ID'] = id.copy()
ticker_symbol = fundamentals_condensed_df.iloc[:,1]
earnings_df['Ticker Symbol'] = ticker_symbol.copy()
for_year = fundamentals_condensed_df.iloc[:,13]
earnings_df['For Year'] = for_year.copy()
earnings_per_share = fundamentals_condensed_df.iloc[:,14]
earnings_df['Earnings Per Share'] = earnings_per_share.copy()
estimated_shares_outstanding = fundamentals_condensed_df.iloc[:,15]
earnings_df['Estimated Earnings'] = estimated_shares_outstanding.copy()

#creating new earnings per share plus minus column
earnings_per_share_plus_minus = []

for value in earnings_df['Earnings Per Share']:
    try:
        if value >= 0:
            earnings_per_share_plus_minus.append('(+)')
        elif value < 0:
            earnings_per_share_plus_minus.append('(-)')
        else:
            earnings_per_share_plus_minus.append('N/A')
    except:
        print('Error. Cannot append plus/minus value(s).')

earnings_df['Earnings Per Share +/-'] = earnings_per_share_plus_minus

#creating new estimated earnings grade column
estimated_earnings_status = []

for value in earnings_df['Estimated Earnings']:
    try:
        if value >= 500000000:
            estimated_earnings_status.append('Excellent')
        elif value >= 25000000 and value <= 499999999.99:
            estimated_earnings_status.append('Solid')
        elif value >= 1 and value <= 249999999.99:
            estimated_earnings_status.append('Positive')
        elif value == 0:
            estimated_earnings_status.append('Even')
        elif value < 0:
            estimated_earnings_status.append('Failing')
        else:
            estimated_earnings_status.append('N/A')
    except:
        print('Error. Cannot append estimated earnings status values.')

earnings_df['Estimated Earnings Grade'] = estimated_earnings_status

# created new workbook containing fundamentals condensed dataframe
earnings_df.to_excel('earnings_df.xlsx', index=False)


# *DATAFRAME PER SELECTED STOCK*

#appl (apple) stock
def appl_yearly_earnings(df):
    return df[df['Ticker Symbol'] == 'AAPL']

aapl_stock_yearly_earnings_per_share_df = \
    appl_yearly_earnings(earnings_df)

aapl_stock_yearly_earnings_per_share_df.\
    to_excel('aapl_stock_yearly_earnings_per_share_df.xlsx', index=False)

#msft (microsoft) stock
def msft_yearly_earnings(df):
    return df[df['Ticker Symbol'] == 'MSFT']

msft_stock_yearly_earnings_per_share_df = \
    msft_yearly_earnings(earnings_df)

msft_stock_yearly_earnings_per_share_df.\
    to_excel('msft_stock_yearly_earnings_per_share_df.xlsx', index=False)

#nflx (netflix) stock
def nflx_yearly_earnings(df):
    return df[df['Ticker Symbol'] == 'NFLX']

nflx_stock_yearly_earnings_per_share_df = \
    nflx_yearly_earnings(earnings_df)

nflx_stock_yearly_earnings_per_share_df.\
    to_excel('nflx_stock_yearly_earnings_per_share_df.xlsx', index=False)

#pfe (pfizer) stock
def pfe_yearly_earnings(df):
    return df[df['Ticker Symbol'] == 'PFE']

pfe_stock_yearly_earnings_per_share_df = \
    pfe_yearly_earnings(earnings_df)

pfe_stock_yearly_earnings_per_share_df.\
    to_excel('pfe_stock_yearly_earnings_per_share_df.xlsx', index=False)

#nke (nike) stock
def nke_yearly_earnings(df):
    return df[df['Ticker Symbol'] == 'NKE']

nke_stock_yearly_earnings_per_share_df = \
    nke_yearly_earnings(earnings_df)

nke_stock_yearly_earnings_per_share_df.\
    to_excel('nke_stock_yearly_earnings_per_share_df.xlsx', index=False)

#bmy (bristol-myers squibb) stock
def bmy_yearly_earnings(df):
    return df[df['Ticker Symbol'] == 'BMY']

bmy_stock_yearly_earnings_per_share_df = \
    bmy_yearly_earnings(earnings_df)

bmy_stock_yearly_earnings_per_share_df.\
    to_excel('bmy_stock_yearly_earnings_per_share_df.xlsx', index=False)


# *REMOVE CERTAIN ROWS UNDER CERTAIN CONDITIONS*

# remove rows from single stock df if the following:
#   earnings per share or estimated earnings is blank
def earnings_per_share_or_estimated_earnings(df):
    return df[df['Earnings Per Share'].notna() |
              df['Estimated Earnings'].notna()]
# nflx
nflx_stock_yearly_earnings_per_share_df = \
    earnings_per_share_or_estimated_earnings(
        nflx_stock_yearly_earnings_per_share_df
    )
# save/updates sheet
nflx_stock_yearly_earnings_per_share_df.\
    to_excel('nflx_stock_yearly_earnings_per_share_df.xlsx', index=False)

# nke
nke_stock_yearly_earnings_per_share_df = \
    earnings_per_share_or_estimated_earnings(
        nke_stock_yearly_earnings_per_share_df
    )
# save/updates sheet
nke_stock_yearly_earnings_per_share_df.\
    to_excel('nke_stock_yearly_earnings_per_share_df.xlsx', index=False)



# *PIVOT TABLES*

#earnings per share pivot table
earnings_per_share_pivot_table = pd.pivot_table(
    earnings_df,
    index='Ticker Symbol',
    columns='For Year',
    values='Earnings Per Share',
    aggfunc='sum'
)

earnings_per_share_pivot_table.to_excel('earnings_per_share_pivot_table.xlsx')


# *MERGING*
#msft and nflx merge
msft_and_nflx_yearly_earnings_per_share = pd.merge(
    msft_stock_yearly_earnings_per_share_df,
    nflx_stock_yearly_earnings_per_share_df,
    on='For Year'
)

msft_and_nflx_yearly_earnings_per_share.\
    to_excel('msft_and_nflx_yearly_earnings_per_share_merge.xlsx', index=False)


# *CONDITIONAL FORMATTING*
#earnings per share pivot table, cells highlighted with certain colors
def highlight_earnings_per_share_pivot_table_values(dollar_amt):
    try:
        if dollar_amt > 0:
            color = '#65F595' # green
        elif dollar_amt == 0:
            color = '#32ADF5' # light blue
        elif dollar_amt < 0:
            color = '#FA291A' # red
        else:
            color = None
        return f'background-color: {color}'
    except:
        print('Error. Unable to highlight cells.')

earnings_per_share_pivot_table = \
    earnings_per_share_pivot_table.style.\
        applymap(highlight_earnings_per_share_pivot_table_values)

earnings_per_share_pivot_table.\
    to_excel('earnings_per_share_pivot_table.xlsx')


# *CHARTS*
# using matplotlib
import matplotlib.pyplot as plt
# AAPL yearly earnings per share pie chart
file = pd.read_excel('aapl_stock_yearly_earnings_per_share_df.xlsx')
plt.title('AAPL Yearly Earnings Per Share')
plt.xlabel("For Year")
plt.ylabel("Earnings Per Share")
plt.pie(file['Earnings Per Share'],labels=file['For Year'])
plt.show()

# MSFT yearly earnings per share pie chart
file = pd.read_excel('msft_stock_yearly_earnings_per_share_df.xlsx')
plt.title('MSFT Yearly Earnings Per Share')
plt.xlabel("For Year")
plt.ylabel("Earnings Per Share")
plt.pie(file['Earnings Per Share'],labels=file['For Year'])
plt.show()

# NFLX yearly earnings per share pie chart
file = pd.read_excel('nflx_stock_yearly_earnings_per_share_df.xlsx')
plt.title('NFLX Yearly Earnings Per Share')
plt.xlabel("For Year")
plt.ylabel("Earnings Per Share")
plt.pie(file['Earnings Per Share'],labels=file['For Year'])
plt.show()

# PFE yearly earnings per share pie chart
file = pd.read_excel('pfe_stock_yearly_earnings_per_share_df.xlsx')
plt.title('PFE Yearly Earnings Per Share')
plt.xlabel("For Year")
plt.ylabel("Earnings Per Share")
plt.pie(file['Earnings Per Share'],labels=file['For Year'])
plt.show()

# NKE yearly earnings per share pie chart
file = pd.read_excel('nke_stock_yearly_earnings_per_share_df.xlsx')
plt.title('NKE Yearly Earnings Per Share')
plt.xlabel("For Year")
plt.ylabel("Earnings Per Share")
plt.pie(file['Earnings Per Share'],labels=file['For Year'])
plt.show()

# BMY yearly earnings per share pie chart
file = pd.read_excel('bmy_stock_yearly_earnings_per_share_df.xlsx')
plt.title('BMY Yearly Earnings Per Share')
plt.xlabel("For Year")
plt.ylabel("Earnings Per Share")
plt.pie(file['Earnings Per Share'],labels=file['For Year'])
plt.show()


# *AUTOMATICALLY ADJUSTING WIDTH FOR ALL COLUMNS*

#using xlwings to auto adjust column width
import xlwings as xw

def auto_fit_excel_columns_and_rows(file_path):
    # open the workbook
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets['Sheet1']  # Adjust as per your sheet name

    # auto-fit column widths and row heights
    sheet.autofit(axis='columns')
    sheet.autofit(axis='rows')

    wb.save()
    app.quit() # close workbook

file_path = 'fundamentals_condensed_df.xlsx'
auto_fit_excel_columns_and_rows(file_path)

file_path = 'earnings_df.xlsx'
auto_fit_excel_columns_and_rows(file_path)

file_path = 'earnings_per_share_pivot_table.xlsx'
auto_fit_excel_columns_and_rows(file_path)