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

if __name__ == '__main__':
    d = json.load(open(r'D:\development\pyNOVATime\pay\2022-06-06--2022-06-19.json'))
    ts = d['DataList'][0]
    df = pd.DataFrame()
    df.insert(0,'key',None,False)
    df.insert(1,'type',None,True)

    for key in ts:
        t, k = strip_type(key)
        parts = inflection.underscore(k).split('_')
        for part in parts:
            if part not in df.columns:
                df.insert(len(df.columns),part,False,True)
        row = dict.fromkeys(parts,True)
        row['type'] = t
        row['key'] = key
        df = df.append(pd.DataFrame(data=[row]))
    print(df)
    df.to_csv(r'D:\development\pyNOVATime\pay\keys.csv')