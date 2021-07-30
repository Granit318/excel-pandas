import pandas as pd

df = pd.read_excel('psmet.xlsx')


def is_table_head(row):
    return (
        not pd.isnull(row.iloc[0])
        and row.iloc[8] == "Цена без НДС"
        and row.iloc[11] == "Цена с НДС"
    )


def read_table(row, price):
    price.append(
        {
            "name": table_header,
            "p1": row.iloc[0],
            "p2": row.iloc[3],
            "price_without_nds": row.iloc[8],
            "price_with_nds": row.iloc[11],
        }
    )


price_data = []
table_found = False
table_header = ""
for index, row in df.iterrows():
    if is_table_head(row):
        table_found = True
        table_header = row.iloc[0]
        continue

    if table_found:
        if not pd.isnull(row.iloc[0]):
            read_table(row, price_data)
        else:
            table_found = False

df = pd.DataFrame(price_data)
df.to_excel('prices.xlsx')
print()