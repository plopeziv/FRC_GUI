import pytest
from unittest.mock import Mock

from data_manager.e_ticket_formats.e_ticket_markup import ETicketMarkup
from tests.stubs.ticket_data_stub import get_ticket_data_stub


@pytest.fixture
def e_ticket_markup_obj():
    ticket_data = get_ticket_data_stub()
    return ETicketMarkup(file_path=".", incoming_ticket=ticket_data, markup=10)


def test_placeholder_e_ticket_markup_imports():
    assert ETicketMarkup is not None


class TestNormalizeLaborToSellPrice:
    def test_positive_markup_normalizes_labor_rates(self, e_ticket_markup_obj):
        e_ticket = e_ticket_markup_obj

        e_ticket._normalize_labor_to_sell_price()

        assert e_ticket.incoming_ticket["Labor"]["RT"]["rate"] == pytest.approx(139.23)
        assert e_ticket.incoming_ticket["Labor"]["OT"]["rate"] == pytest.approx(179.73)
        assert e_ticket.incoming_ticket["Labor"]["DT"]["rate"] == pytest.approx(216.11)
        assert e_ticket.incoming_ticket["Labor"]["OT DIFF"]["rate"] == pytest.approx(36.86)
        assert e_ticket.incoming_ticket["Labor"]["DT DIFF"]["rate"] == pytest.approx(73.25)

    def test_zero_markup_leaves_labor_rates_unchanged(self, e_ticket_markup_obj):
        e_ticket = e_ticket_markup_obj
        e_ticket.markup = 0

        e_ticket._normalize_labor_to_sell_price()

        assert e_ticket.incoming_ticket["Labor"]["RT"]["rate"] == "153.15"
        assert e_ticket.incoming_ticket["Labor"]["OT"]["rate"] == "197.70"
        assert e_ticket.incoming_ticket["Labor"]["DT"]["rate"] == "237.72"
        assert e_ticket.incoming_ticket["Labor"]["OT DIFF"]["rate"] == "40.55"
        assert e_ticket.incoming_ticket["Labor"]["DT DIFF"]["rate"] == "80.57"
