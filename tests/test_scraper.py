import pytest
from tax_scraper import strip_street_suffix, parse_pct

def test_strip_street_suffix():
    # Test common suffixes are stripped
    assert strip_street_suffix("500 W Madison St") == "500 W Madison"
    assert strip_street_suffix("100 S Wacker Dr") == "100 S Wacker"
    assert strip_street_suffix("123 Main Avenue") == "123 Main"
    assert strip_street_suffix("456 Oak Road") == "456 Oak"
    
    # Test trailing periods are handled
    assert strip_street_suffix("789 Pine Lane.") == "789 Pine"
    
    # Test case insensitivity of suffixes
    assert strip_street_suffix("321 Maple COURT") == "321 Maple"
    
    # Test address without suffix remains untouched
    assert strip_street_suffix("100 Broadway") == "100 Broadway"
    
    # Test single word address
    assert strip_street_suffix("Loop") == "Loop"

def test_parse_pct():
    # Test standard percentages
    assert parse_pct("1.750%") == 0.0175
    assert parse_pct("10.25%") == 0.1025
    assert parse_pct("0.00%") == 0.0
    
    # Test whitespace handling
    assert parse_pct("  6.25%  ") == 0.0625
    
    # Test fallback on invalid strings
    assert parse_pct("invalid") == 0.0
    assert parse_pct(None) == 0.0
