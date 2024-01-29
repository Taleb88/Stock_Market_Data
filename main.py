import pandas as pd

fundamentals_df = pd.read_excel('master.xlsx',
                                sheet_name="fundamentals")
prices_df = pd.read_excel('master.xlsx',
                          sheet_name="prices")
prices_split_adjusted_df = pd.read_excel('master.xlsx',
                                         sheet_name="prices-split-adjusted")
securities_df = pd.read_excel('master.xlsx',
                              sheet_name="securities")

print(fundamentals_df.head())
print(fundamentals_df.tail())