import pytest
import os

from datetime import datetime
from os.path import join, dirname, realpath

from ExcelReader import DictConverter, Field

def test_load_file_horizontal_data():
    __location__ = realpath(
        join(os.getcwd(), dirname(__file__))
    )
    r = DictConverter(join(__location__,'simple_test_values.xlsx'),
               Field(0, 0),
               Field(0, 3)
    )
    assert r, "The reader couldn't be created."
    assert r.get_headers() == ['header1', 'header2', 'header3', 'header4'], 'The headers were incorrectly loaded.'
    data = r.get_data()

    assert len(data) == 4, 'Expected 4 elements at the horizontal data sheet.'
    assert (data[3]['header1'] == True and
            data[3]['header2'] == datetime(2019, 6, 26) and
            data[3]['header3'] == 'value10' and
            data[3]['header4'] == 20
    ), 'The expected values for the 4th object at the horizontal sheet do not match.'



def test_load_file_vertical_data():
    __location__ = realpath(
        join(os.getcwd(), dirname(__file__))
    )
    r = DictConverter(join(__location__,'simple_test_values.xlsx'),
               Field(0, 0),
               Field(3, 0),
               sheet=1
    )
    assert r, "The reader couldn't be created."
    assert r.get_headers() == ['header11', 'header22', 'header33', 'header44'], 'The headers were incorrectly loaded.'
    data = r.get_data()
    assert len(data) == 4, 'Expected 4 elements at the vertical data sheet.'
    assert (data[3]['header11'] == False and
            data[3]['header22'] == datetime(2019, 6, 26) and
            data[3]['header33'] == 'value19' and
            data[3]['header44'] == 32
    ), 'The expected values for the 4th object at the vertical sheet do not match.'
