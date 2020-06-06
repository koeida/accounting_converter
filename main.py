import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import sys

from collections import namedtuple

Row = namedtuple("Row", ["date", "account", "amount", "explanation", "name"])

def read_sheet(f, sheet):
    df = pd.read_excel(f, sheet)
    journal_entry_name = df[df.columns[0]][2]
    rows = []
    for index, row in df.iterrows():
        date, acct = row[:2]
        if len(row) >= 6:
            amount, explanation, name = row[3:6]
        else:
            amount, explanation = row[3:5]
            name = ""
        rows.append(Row(date, acct, amount, explanation, name))
    return (journal_entry_name, rows)

def row_to_str(journal_entry_name, r):
    jname_str = '{:7.7}'.format(journal_entry_name)
    dstr = r.date.strftime("%m%d%y")
    acct_str = '{:10.10}'.format(str(r.account))
    amount_str = '{:12.12}'.format(str(round(r.amount,2)))
    if str(r.explanation) == "nan":
        memo_str = '{:40.40}'.format("%s" % r.name)
    elif str(r.name) == "nan":
        memo_str = '{:40.40}'.format("%s" % r.explanation)
    else:
        memo_str = '{:40.40}'.format("%s: %s" % (r.explanation, r.name))
    return "%s%s%s%s%s" % (jname_str, dstr, acct_str, amount_str, memo_str)

def valid_row(r):
    return str(r.date) != "nan"

def partition(f,l):
    yes = list(filter(f,l))
    no = list(filter(lambda x: not f(x),l))
    return (yes,no)

js = pd.read_excel(sys.argv[1], None)
final_results = []
final_errors = []
for jname in js:
    journal_entry_name, rows = read_sheet(sys.argv[1], jname)
    rows = rows[4:]
    results, errors = partition(valid_row, rows)
    results = list(map(lambda r: row_to_str(journal_entry_name, r), results))
    final_results.extend(results)
    final_errors.extend(errors)
    

f = open("output.txt", "w")
for r in final_results:
    print(r, file=f)

for r in final_errors:
    print("Error on row: %s" % str(r), file=f)
f.close()


