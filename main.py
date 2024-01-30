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

# remove timestamp from period ending values
fundamentals_condensed_df['Period Ending'] = \
    fundamentals_condensed_df['Period Ending'].astype(str).str[:11] #convert to string format before slicing value

# created new workbook containing fundamentals condensed dataframe
fundamentals_condensed_df.to_excel('fundamentals_condensed_df.xlsx', index=False)