import pytest
from pst_kb.language import detect_language

def test_detect_language_hebrew():
    # Enough Hebrew characters should trigger 'he'
    text = "שלום מה שלומך היום? אני כותב לך הודעה בעברית."
    assert detect_language(text) == "he"

def test_detect_language_english():
    text = "Hello, how are you today? I am writing you a message in English."
    assert detect_language(text) == "en"

def test_detect_language_empty():
    assert detect_language("") is None
    assert detect_language("   ") is None

def test_detect_language_mixed():
    # Mixed text, should prefer langdetect or dominant
    text = "Hello שלום"
    # Depending on langdetect's behavior with short strings
    result = detect_language(text)
    assert result in ["he", "en", None]
