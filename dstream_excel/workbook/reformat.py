import re
import pandas as pd
from morningstar.exceptions import DataNotPopulatedException


def reformat_data(df):
    _id, id_type = _get_etf_id_and_id_type_from_holding_command_in_df(df)
    df = _transpose_remove_extra_and_rename_columns(df)
    return _reformat_data(df, _id, id_type)

##### Helper functions below #####

def _reformat_data(df, _id, id_type):
    # Go through each column and form df
    out_df = pd.DataFrame(columns=['Date','MV','TICKER','ETF TICKER','CUSIP','ETF CUSIP'])
    for i in range(1, len(df.columns)):
        col = df.iloc[:, i]
        out_df = out_df.append(_left_column_plus_column_as_df_where_second_col_is_orig_colname(
            df['Date'],
            col,
            'MV',
            id_type
        ))

    # Reset index
    out_df.reset_index(drop=True, inplace=True)

    # Add etf info to all data
    out_df['ETF ' + id_type.upper()] = _id

    #Reorder columns
    out_df = out_df[['Date','MV','TICKER','ETF TICKER','CUSIP','ETF CUSIP']]

    return out_df


def _left_column_plus_column_as_df_where_second_col_is_orig_colname(left_col, col, colname_orig, colname_new):
    right = _column_as_df_where_second_col_is_orig_colname(col, colname_orig, colname_new)
    return pd.concat([left_col, right], axis=1)


def _column_as_df_where_second_col_is_orig_colname(col, colname_orig, colname_new):
    new_col_value = col.name
    df = pd.DataFrame(col)
    df.rename(columns={new_col_value: colname_orig}, inplace=True)
    df[colname_new] = new_col_value
    return df


def _transpose_remove_extra_and_rename_columns(df):
    df = df.T.drop(1, axis=0)
    df.iloc[0, 0] = 'Date'
    df.columns = df.iloc[0]
    df = df.drop(0, axis=0).reset_index(drop=True)
    return df


def _get_etf_id_and_id_type_from_holding_command_in_df(df):
    holding_formula_value = df.iloc[0, 0]

    p = re.compile(r'=_xll\.MSHOLDING\("(\w*)","(\w*)"[\S\s]*\)')
    m = p.match(holding_formula_value)

    if not m:
        raise DataNotPopulatedException(f'Could not parse formula {holding_formula_value}. ' + \
                                        'This means that data was not populated correctly from Morningstar.')

    return m.group(1), m.group(2)