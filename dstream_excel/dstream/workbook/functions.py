from typing import Sequence, Union

class DatastreamExcelFunction:
    def __init__(self, symbols: Union[Sequence[str], str]):
        self.symbols = symbols

    def time_series(self, variables: Sequence[str], begin='-2Y', end='', freq='D'):
        return _datastream_excel_function(self.symbols, variables, begin=begin, end=end, freq=freq)

    def static(self, variables: Sequence[str]):
        return _datastream_excel_function(self.symbols, variables, begin='Latest Value', end='', freq='')


def _datastream_excel_function(symbols: Union[Sequence[str], str], variables: Sequence[str], begin, end, freq):
    variables_str = ';'.join(variables)
    if isinstance(symbols, (list, tuple)):
        symbols = ','.join(symbols)
    middle_str = '","'.join([symbols, variables_str, begin, end, freq])
    custom_header_str = 'CustomHeader=true;CustHeaderDatatypes=DATATYPE'
    headers_str = f'RowHeader=true;ColHeader=false'
    other_str = 'DispSeriesDescription=false;YearlyTSFormat=false;QuarterlyTSFormat=false'
    return f'=DSGRID("{middle_str}","{custom_header_str};{headers_str};{other_str}","")'