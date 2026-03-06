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
