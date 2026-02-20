from data_manager.e_ticket_creator import ETicketCreator
from data_manager.e_ticket_formats.e_ticket_markup import ETicketMarkup

TEMPLATES = {
    "STANDARD TEMPLATE": {
        "allow_markup": False,
        "default_markup": None,
        "eTicket_lambda": lambda folder_path, data: ETicketCreator(folder_path, data)
        },
    "CHASE MARKUP TEMPLATE": {
        "allow_markup": True,
        "default_markup": 10,
        "eTicket_lambda": lambda folder_path, data, **kwargs: ETicketMarkup(folder_path, data, **kwargs)
        }
}