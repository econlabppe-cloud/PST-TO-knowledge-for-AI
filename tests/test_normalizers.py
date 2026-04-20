import pytest
from pst_kb.normalizers.text import (
    normalize_whitespace,
    normalize_subject,
    normalize_email,
    detect_reply_forward_indicator,
)

def test_normalize_whitespace():
    assert normalize_whitespace("  hello  world  ") == "hello world"
    assert normalize_whitespace("line1\r\nline2") == "line1\nline2"
    assert normalize_whitespace("line1\n\n\nline2") == "line1\n\nline2"

def test_normalize_subject():
    assert normalize_subject("Re: hello") == "hello"
    assert normalize_subject("Fwd: [EXTERNAL] meeting") == "meeting"
    assert normalize_subject("תגובה: שלום") == "שלום"
    assert normalize_subject("FWD: RE: FW: test") == "test"

def test_normalize_email():
    assert normalize_email(" <USER@example.com> ") == "user@example.com"
    assert normalize_email(None) is None
    assert normalize_email("  ") is None

def test_detect_reply_forward_indicator():
    assert detect_reply_forward_indicator("Re: topic") == "reply"
    assert detect_reply_forward_indicator("Fwd: topic") == "forward"
    assert detect_reply_forward_indicator("תגובה: נושא") == "reply"
    assert detect_reply_forward_indicator("הועבר: נושא") == "forward"
    assert detect_reply_forward_indicator("Hello") is None
