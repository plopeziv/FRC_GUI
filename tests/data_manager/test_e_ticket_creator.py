import pytest

from data_manager.e_ticket_creator import ETicketCreator
from config.equipment_lists import EXCLUDED_EQUIPMENT
from tests.stubs.ticket_data_stub import get_ticket_data_stub


class TestSplitMaterialsAndEquipment:
    def test_splits_materials_using_config_excluded_equipment(self):
        ticket_data = get_ticket_data_stub()
        grouped_materials = ticket_data["Materials"]

        materials, equipment = ETicketCreator._split_materials_and_equipment(
            grouped_materials, EXCLUDED_EQUIPMENT
        )

        assert [item["material"] for item in materials] == [
            "MAPEI PLANIPREP SC 10LB BAG",
            "MAPEI QUICK PATCH 25LB",
        ]
        assert [item["material"] for item in equipment] == [
            "HEPA SANDER#302 & VAC #701",
            "TURBO STRIPPER # 203",
        ]

    def test_returns_empty_lists_when_grouped_materials_is_empty(self):
        materials, equipment = ETicketCreator._split_materials_and_equipment(
            [], EXCLUDED_EQUIPMENT
        )

        assert materials == []
        assert equipment == []


class TestInsertItemsAtAnchor:
    def test_raises_value_error_when_anchor_is_not_found(self, monkeypatch):
        creator = ETicketCreator()
        ws = object()

        monkeypatch.setattr(creator, "_find_material_row", lambda *_args, **_kwargs: None)

        with pytest.raises(ValueError, match="Material Used starting row not found"):
            creator._insert_items_at_anchor(
                ws,
                items=[{"material": "A", "quantity": "1", "units": "EA", "sell price": "1"}],
                anchor_search_term="Material Used",
                anchor_column="A",
                price_key="sell price",
            )

    def test_inserts_items_starting_one_row_after_anchor(self, monkeypatch):
        creator = ETicketCreator()
        ws = object()
        items = [{"material": "A", "quantity": "1", "units": "EA", "sell price": "1"}]
        calls = {}

        monkeypatch.setattr(creator, "_find_material_row", lambda *_args, **_kwargs: 10)

        def fake_insert_line_items(ws_arg, items_arg, start_row_arg, price_key_arg):
            calls["ws"] = ws_arg
            calls["items"] = items_arg
            calls["start_row"] = start_row_arg
            calls["price_key"] = price_key_arg
            return 14

        monkeypatch.setattr(creator, "_insert_line_items", fake_insert_line_items)

        last_row = creator._insert_items_at_anchor(
            ws,
            items=items,
            anchor_search_term="Material Used",
            anchor_column="A",
            price_key="sell price",
        )

        assert last_row == 14
        assert calls["ws"] is ws
        assert calls["items"] == items
        assert calls["start_row"] == 11
        assert calls["price_key"] == "sell price"

    def test_returns_anchor_plus_one_for_empty_items(self, monkeypatch):
        creator = ETicketCreator()
        ws = object()

        monkeypatch.setattr(creator, "_find_material_row", lambda *_args, **_kwargs: 10)
        monkeypatch.setattr(creator, "_insert_line_items", lambda *_args, **_kwargs: 11)

        last_row = creator._insert_items_at_anchor(
            ws,
            items=[],
            anchor_search_term="Material Used",
            anchor_column="A",
            price_key="sell price",
        )

        assert last_row == 11
