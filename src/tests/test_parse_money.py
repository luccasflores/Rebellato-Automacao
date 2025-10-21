import math
from src.core.utils import parse_money

def test_parse_money_simple():
    assert parse_money("10") == 10.0
    assert parse_money("10.50") == 10.5
    assert parse_money("10,50") == 10.5
    assert parse_money("1.234,56") == 1234.56
    assert parse_money("1,234.56") == 1234.56
    assert math.isnan(parse_money(None))
