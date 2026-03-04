import pytest

from data_manager.e_ticket_formats.e_ticket_markup import ETicketMarkup


def test_placeholder_e_ticket_markup_imports():
    assert ETicketMarkup is not None
