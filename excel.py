import pandas as pd

df = pd.read_excel("cvetmet.xls")


def is_table_head(row):
    return (
        not pd.isnull(row.iloc[0])
        and row.iloc[3] == "Цена"
    )


def read_table(row, price):
    price.append(
        {
            "name": table_header,
            "марка": row.iloc[0],
            "p2": row.iloc[1],
            "ед.изм": row.iloc[2],
            "price": row.iloc[3],

        }
    )


price_data = []
table_found = False
table_header = ""
for index, row in df.iterrows():
    if is_table_head(row):
        table_found = True
        table_header = df.iloc[index - 1].iloc[0]
        continue

    if table_found:
        if not pd.isnull(row.iloc[0]):
            read_table(row, price_data)
        else:
            table_found = False
result = pd.DataFrame(price_data)
# df_old = pd.read_excel('prices.xlsx')
# df = df.append(df_old, ignore_index=True)
result.to_excel('prices.xlsx')
print()



