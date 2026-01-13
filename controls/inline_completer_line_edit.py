from qtpy.QtWidgets import QLineEdit, QCompleter
from qtpy.QtCore import Qt

class InlineCompleterLineEdit(QLineEdit):
    def __init__(self, items, parent=None):
        super().__init__(parent)

        # Create a completer with the list of items
        self.completer = QCompleter(items, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)  # match anywhere
        self.completer.setCompletionMode(QCompleter.InlineCompletion)  # <— key line for inline autocomplete

        # Connect the completer to this QLineEdit
        self.setCompleter(self.completer)