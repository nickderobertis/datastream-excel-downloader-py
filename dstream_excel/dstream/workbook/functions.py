from typing import Sequence

class DatastreamExcelFunction:
    def __init__(self, symbols):
        self.symbols = symbols

    def time_series(self, variables, begin='-2Y', end='', freq='D', headers=True):
        return _datastream_excel_function(self.symbols, variables, begin=begin, end=end, freq=freq, headers=headers)

    def static(self, variables, headers=True):
        return _datastream_excel_function(self.symbols, variables, begin='Latest Value', end='', freq='',
                                          headers=headers)


def _datastream_excel_function(symbols, variables: Sequence[str], begin, end, freq, headers):
    variables_str = ';'.join(variables)
    middle_str = '","'.join([symbols, variables_str, begin, end, freq])
    headers_bool_str = str(headers).lower()
    headers_str = f'RowHeader={headers_bool_str};ColHeader={headers_bool_str}'
    other_str = 'DispSeriesDescription=false;YearlyTSFormat=false;QuarterlyTSFormat=false'
    return f'=DSGRID("{middle_str}","{headers_str};{other_str}","")'