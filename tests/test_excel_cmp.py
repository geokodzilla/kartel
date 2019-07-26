# -*- coding: utf-8 -*-

import pytest
from kartel.excel_cmp import ExcelCmp
from collections import OrderedDict

@pytest.fixture
def xlsx():
    old_data = OrderedDict([('5-100;1', ['1', '5-100;1', '1110', 'Dj']),
                            ('5-100;2', ['2', '5-100;2', '1271', 'In']),
							('5-100;3', ['3', '5-100;3', '1271', 'In']),
							('5-101;1', ['4', '5-101;1', '1110', 'Dj'])])
    new_data = OrderedDict([('5-100;1', ['1', '5-100;1', '1110', 'Dj']),
                            ('5-100;2', ['2', '5-100;2', '1271', 'Bg']),
							('5-101;1', ['4', '5-101;1', '1110', 'Dj']),
                            ('5-101;2', ['5', '5-101;2', '1271', 'In'])])
    return ExcelCmp(old_data, new_data)


def test_excel_cmp(xlsx):
    xlsx.cmp()
    for row in xlsx.mod:
        assert row in [('M', ['2',	'5-100;2',	'1271',	'In -> Bg']),
                    ('U', ['3',	'5-100;3',	'1271', 'In']),
                    ('D', ['5',	'5-101;2', 	'1271', 'In'])]
                    
                    
def test_clear_mod(xlsx):
    xlsx.cmp()
    assert xlsx.mod != []
    xlsx.clear_mod()
    assert xlsx.mod == []