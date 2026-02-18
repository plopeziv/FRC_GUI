TEMPLATES = {
    "STANDARD TEMPLATE": lambda folder_path, data: ETicketCreator(folder_path, data),
    "CHASE MARKUP TEMPLATE": lambda folder_path, data, **kwargs: ETicketMarkup(folder_path, data, **kwargs),
    "MATERIAL MARKUP TEMPLATE": ""
}