import pytest

from fibooks.excel_parser import create_dataset, get_accounts

def test_create_dataset_with_valid_data():
    data = {
        0 : ['A', 'Date', 'revenue', 'net profit'],
        1 : ['B', '2023-09-30',100, 80],
        2 : ['C', '2023-06-30',10, 8],
    }
    accounts = get_accounts(data)
    result = create_dataset(data, accounts)
    assert result == {'date': ['2023-09-30','2023-06-30'], 'net profit': [80, 8], 'revenue': [100, 10]}


def test_create_dataset_with_valid_data_with_empty_rows():
    data = {
        0 : ['A', None, 'Date', None, 'revenue', 'net profit'],
        1 : ['B', None, '2023-09-30', None,100, 80],
        2 : ['C', None, '2023-06-30', None,10, 8],
    }
    accounts = get_accounts(data)
    result = create_dataset(data, accounts)
    assert result == {'date': ['2023-09-30','2023-06-30'], 'net profit': [80, 8], 'revenue': [100, 10]}

def test_create_dataset_with_empty_data():
    data = {}
    accounts = []
    result = create_dataset(data, accounts)
    assert result == {}