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

fundamentals_condensed_df = pd.DataFrame()
id = fundamentals_df.iloc[:,0]
fundamentals_condensed_df['id'] = id.copy()

fundamentals_condensed_df.to_excel('fundamentals_condensed_df.xlsx', index=False)