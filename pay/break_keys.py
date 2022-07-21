import inflection
import json
import pandas as pd


def strip_type(key):
    if key[0] == key[0].lower():
        t = key[0]
        key = key[1:]
    else:
        t = ''
    return (t, key)


def make_df(tags):
    df = pd.DataFrame()
    df.insert(0, 'PYTHON_KEY', None, False)
    df.insert(1, 'PYTHON_TYPE', None, True)
    df.insert(2, 'PYTHON_CLASS', None, True)

    for key in tags:
        t, k = strip_type(key)
        parts = inflection.underscore(k).split('_')
        for part in parts:
            if part not in df.columns:
                df.insert(len(df.columns), part, pd.NA, True)
        row = dict.fromkeys(parts, True)
        row['PYTHON_TYPE'] = t
        row['PYTHON_KEY'] = key
        df = df.append(pd.DataFrame(data=[row]))
    return df


def reorder_columns(df):
    new_columns = []
    for column in df.columns:
        if 'PYTHON' not in column:
            try:
                new_columns.append((column, df[column].value_counts()[True]))
            except KeyError:
                new_columns.append((column, 0))
    new_columns = [tag[0]
                   for tag in sorted(sorted(new_columns), reverse=True, key=lambda tup: tup[1])]
    return df.reindex(columns=['PYTHON_KEY', 'PYTHON_TYPE', 'PYTHON_CLASS']+new_columns)


if __name__ == '__main__':
    d = json.load(
        open(r'D:\development\pyNOVATime\pay\2022-06-06--2022-06-19.json'))
    ts = d['DataList'][0]
    df = make_df(ts)
    df = reorder_columns(df)
    df.to_csv(r'D:\development\pyNOVATime\pay\keys.csv')
