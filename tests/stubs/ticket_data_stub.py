"""Reusable ticket-data stubs for tests."""

from copy import deepcopy


TICKET_DATA_STUB = {
    "Job Number": "123456",
    "Job Name": "Test Job",
    "Ticket Number": "00003",
    "Job Address": "1234 Fake Street",
    "Date": "11/10/25",
    "Signature": "Yes",
    "Type": "REGULAR",
    "Installers": "Juan",
    "Work Location": "35TH FLR",
    "Description": "Fixed pump housing leak",
    "Labor": {
        "RT": {"hours": "12", "rate": "153.15"},
        "OT": {"hours": "4", "rate": "197.70"},
        "DT": {"hours": "0", "rate": "237.72"},
        "OT DIFF": {"hours": "0", "rate": "40.55"},
        "DT DIFF": {"hours": "0", "rate": "80.57"},
    },
    "Materials": [
        {
            "material": "MAPEI PLANIPREP SC 10LB BAG",
            "quantity": "3",
            "units": "BG",
            "unit cost": "23.55",
            "sell price": "34.35",
        },
        {
            "material": "MAPEI QUICK PATCH 25LB",
            "quantity": "10",
            "units": "BG",
            "unit cost": "23.97",
            "sell price": "34.87",
        },
        {
            "material": "HEPA SANDER#302 & VAC #701",
            "quantity": "2",
            "units": "EA",
            "unit cost": "150",
            "sell price": "150",
        },
        {
            "material": "TURBO STRIPPER # 203",
            "quantity": "1",
            "units": "EA",
            "unit cost": "315",
            "sell price": "315",
        },
    ],
}


def get_ticket_data_stub():
    """Return a deep copy so tests can mutate without cross-test leakage."""
    return deepcopy(TICKET_DATA_STUB)
