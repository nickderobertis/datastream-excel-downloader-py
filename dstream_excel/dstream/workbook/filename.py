

replacements = {
        ':': 'colon',
        '*': 'asterisk'
    }

reverse_replacements = {value: key for key, value in replacements.items()}

def _valid_filename_from_symbol(symbol):
    return _replace_str_with_dict(symbol, replacements)

def _original_symbol_from_filename_symbol(symbol):
    return _replace_str_with_dict(symbol, reverse_replacements)

def _replace_str_with_dict(symbol, dict_):
    for replacement in dict_:
        symbol = str(symbol).replace(replacement, dict_[replacement])

    return symbol