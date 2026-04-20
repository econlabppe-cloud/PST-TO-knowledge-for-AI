import pytest
from pst_kb.cleaners.email_cleaner import EmailCleaner

def test_email_cleaner_basic():
    cleaner = EmailCleaner()
    text = "Hello world\n\nBest regards,\nJohn"
    result = cleaner.clean(text)
    assert "Hello world" in result.text
    assert "Best regards" not in result.text
    assert "signature_removed" in result.notes

def test_email_cleaner_hebrew_signature():
    cleaner = EmailCleaner()
    text = "שלום רב\n\nבברכה,\nישראל ישראלי"
    result = cleaner.clean(text)
    assert "שלום רב" in result.text
    assert "בברכה" not in result.text
    assert "signature_removed" in result.notes

def test_email_cleaner_reply_cut():
    cleaner = EmailCleaner()
    text = "New message\n\n-----Original Message-----\nFrom: sender\nSent: date\nTo: recipient\nSubject: sub\n\nOld message"
    result = cleaner.clean(text)
    assert "New message" in result.text
    assert "Old message" not in result.text
    assert "reply_history_removed" in result.notes

def test_email_cleaner_mostly_quoted():
    cleaner = EmailCleaner()
    text = "> quoted line 1\n> quoted line 2\nnew line"
    result = cleaner.clean(text)
    assert "new line" in result.text
    assert result.mostly_quoted is True
