from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QAbstractItemView
from PySide6.QtGui import QKeySequence, QGuiApplication
from PySide6.QtCore import Qt, QEvent

class CustomTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed | QAbstractItemView.EditTrigger.AnyKeyPressed)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def keyPressEvent(self, event: QEvent):
        if event.matches(QKeySequence.StandardKey.Paste):
            self.custom_paste()
        elif event.key() == Qt.Key.Key_Delete or event.key() == Qt.Key.Key_Backspace:
            self.delete_selected_cells_content()
        else:
            super().keyPressEvent(event)

    def delete_selected_cells_content(self):
        for item in self.selectedItems():
            item.setText("")

    def custom_paste(self):
        clipboard = QGuiApplication.clipboard()
        mime_data = clipboard.mimeData()

        if mime_data.hasText():
            text = mime_data.text()
            rows_data = text.strip('\n').split('\n')
            if not rows_data: return

            table_data = []
            for row_str in rows_data:
                table_data.append(row_str.split('\t'))

            start_row = self.currentRow() if self.currentRow() != -1 else 0
            start_col = self.currentColumn() if self.currentColumn() != -1 else 0
            if not self.selectedIndexes():
                start_row, start_col = 0, 0

            num_pasted_rows = len(table_data)
            max_pasted_cols_in_data = 0
            if table_data:
                max_pasted_cols_in_data = max(len(row_content) for row_content in table_data) if table_data[0] else 0

            required_rows = start_row + num_pasted_rows
            if required_rows > self.rowCount():
                self.setRowCount(required_rows)

            for r_idx, row_content in enumerate(table_data):
                current_table_row = start_row + r_idx
                for c_idx, cell_value in enumerate(row_content):
                    current_table_col = start_col + c_idx
                    if current_table_col < self.columnCount():
                        item = self.item(current_table_row, current_table_col)
                        if not item:
                            item = QTableWidgetItem(cell_value)
                            self.setItem(current_table_row, current_table_col, item)
                        else:
                            item.setText(cell_value)
        else:
            paste_event = QEvent(QEvent.Type.KeyPress)
            super().keyPressEvent(paste_event)