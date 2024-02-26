import pandas

converters = {
    'Home Team': str,
    'Away Team': str,
    'Date': str,
    'Bet': str,
    'Stake': float,
    'Odds': float,
    'Outcome': str,
    'Sport': str,
    'Profit': float
}

df = pandas.read_excel('./data/Bets_to_analyse.xlsx', converters=converters, sheet_name="Data")
df.dropna(subset=["Home Team", "Away Team", "Date"], inplace=True)

rows = {}

# Filter unique rows and combine values into intermediate dictionary
for index, row in df.iterrows():
    if row["Home Team"] == '':
        continue
    # Find distinct values in "Home Team", "Away Team" and "Date" columns
    _identifier = row["Home Team"] + row["Away Team"] + row["Date"]
    rows[_identifier] = rows.get(_identifier, {
        "Home Team": row["Home Team"],
        "Away Team": row["Away Team"],
        "Date": row["Date"],
        "Bets": row["Bet"],
        "Stakes": [],
        "Profits": [],
        "Odds_List": [],
        "Outcomes": [],
        "Sport": row["Sport"]
    })

    rows[_identifier]["Stakes"].append(row["Stake"])
    rows[_identifier]["Profits"].append(row["Profit"])
    rows[_identifier]["Odds_List"].append(row["Odds"])
    rows[_identifier]["Outcomes"].append(row["Outcome"])

# Calculate the final values
for index, row in rows.items():
    row["Stake"] = round(sum(row["Stakes"]), 2)
    row["Profit"] = round(sum(row["Profits"]), 2)

    row["Odds"] = round(sum([x * y for x, y in zip(row["Stakes"], row["Odds_List"])]) / sum(row["Stakes"]) if sum(row["Stakes"]) != 0 else 1, 3)

    # stakes = sum(row["Stakes"]) if sum(row["Stakes"]) != 0 else 1
    # odds_amount = 0
    # for i in range(len(row["Stakes"])):
    #     odds_amount += row["Stakes"][i] * row["Odds_List"][i]
    # row["Odds"] = round(odds_amount / stakes, 3)

    row['Outcome'] = 'L' if row['Profit'] == -row['Profit'] else 'HL' if row['Profit'] < 0 else 'V' if row['Profit'] == 0 else 'W' if row['Stake'] * (row['Odds'] - 1) == row['Profit'] else 'HW'

    # if row['Profit'] == -row['Profit']:
    #     row['Outcome'] = 'L'
    # elif row['Profit'] < 0:
    #     row['Outcome'] = 'HL'
    # elif row['Profit'] == 0:
    #     row['Outcome'] = 'V'
    # elif row['Stake'] * (row['Odds'] - 1) == row['Profit']:
    #     row['Outcome'] = 'W'
    # else:
    #     row['Outcome'] = 'HW'

    row.pop("Stakes", None)
    row.pop("Profits", None)
    row.pop("Odds_List", None)
    row.pop("Outcomes", None)

df = pandas.DataFrame(rows.values())
print(df.head())

# Engine openpyxl is used to append to the existing file
# if_sheet_exists="replace" is used to replace the sheet if it already exists
# mode="a" is used to append to the existing file
with pandas.ExcelWriter("./data/Bets_to_analyse.xlsx", engine='openpyxl', mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="Distinct Values")
