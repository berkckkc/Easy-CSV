import csv
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidget, QTableWidgetItem, QVBoxLayout, QPushButton, QFileDialog, QAction, QMenu, QWidget, QUndoStack , QUndoCommand , QShortcut , QMessageBox  
from PyQt5.QtGui import  QKeySequence
from PyQt5.QtCore import Qt
import pandas as pd
### Easy CSV  v1.0 --- by Berk ÇIKIKCI --- don't forget to follow on Linkedin for further projects ----  ###
class SetItemCommand(QUndoCommand):
    def __init__(self, table, row, col, prev_item, new_item):
        super().__init__()
        self.table = table
        self.row = row
        self.col = col
        self.prev_text = prev_item.text() if prev_item else None
        self.new_text = new_item.text() if new_item else None

    def undo(self):
        if self.prev_text is not None:
            self.table.setItem(self.row, self.col, QTableWidgetItem(self.prev_text))
        else:
            self.table.takeItem(self.row, self.col)

    def redo(self):
        if self.new_text is not None:
            self.table.setItem(self.row, self.col, QTableWidgetItem(self.new_text))
        else:
            self.table.takeItem(self.row, self.col)

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Easy CSV'
        self.left = 0
        self.top = 0
        self.width = 800
        self.height = 600
        self.deleted_cells = []
        self.csv_loaded = False
        self.undoStack = QUndoStack(self)

        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.layout = QVBoxLayout()
        deleteShortcut = QShortcut(QKeySequence.Delete, self)
        deleteShortcut.activated.connect(self.delete)
        undoShortcut = QShortcut(QKeySequence("Ctrl+Z"), self)
        undoShortcut.activated.connect(self.undo)
        

        self.tableWidget = QTableWidget()
        self.tableWidget.setRowCount(10000)
        self.tableWidget.setColumnCount(10000)
        headers = []
        for i in range(10000):
            headers.append(f"C{i + 1}")
        self.tableWidget.setHorizontalHeaderLabels(headers)
        row_labels = [f"R{i + 1}" for i in range(10000)]
        self.tableWidget.setVerticalHeaderLabels(row_labels)

        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.contextMenuRequested)

        self.layout.addWidget(self.tableWidget)

        self.importButton = QPushButton("Import CSV")
        self.importButton.clicked.connect(self.importCSV)
        self.layout.addWidget(self.importButton)

        self.exportButton = QPushButton("Export CSV")
        self.exportButton.clicked.connect(self.exportCSV)
        self.layout.addWidget(self.exportButton)

        self.mainWidget = QWidget()
        self.mainWidget.setLayout(self.layout)
        self.setCentralWidget(self.mainWidget)


    def load_csv(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv)", options=options)

        if file_name:
            try:
                self.csv_data = pd.read_csv(file_name)
                self.csv_loaded = True
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
        

        

    def contextMenuRequested(self, position):
        contextMenu = QMenu(self)
        copy_action = QAction("Copy", self)
        copy_action.triggered.connect(self.copy)
        contextMenu.addAction(copy_action)
        
        paste_action = QAction("Paste", self)
        paste_action.triggered.connect(self.paste)
        contextMenu.addAction(paste_action)

        cut_action = QAction("Cut", self)
        cut_action.triggered.connect(self.cut)
        contextMenu.addAction(cut_action)

        undo_action = QAction("Undo", self)
        undo_action.triggered.connect(self.undoStack.undo)
        contextMenu.addAction(undo_action)

        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(self.delete)
        contextMenu.addAction(delete_action)

        contextMenu.exec_(self.tableWidget.mapToGlobal(position))

    def keyPressEvent(self, event):
        selected = self.tableWidget.selectedRanges()
        if not selected:
            return
        
        if event.key() == Qt.Key_Backspace:
            self.delete()
        elif event.matches(QKeySequence.Copy):
            self.copy()
        elif event.matches(QKeySequence.Paste):
            self.paste()
        elif event.matches(QKeySequence.Cut):
            self.cut()
        elif event.matches(QKeySequence.Undo):
            self.undoStack.undo()


    def delete(self):
        self.tableWidget.setUpdatesEnabled(False)
        
        # Deleted cells temp
        temp_deleted = []
        
        selected_range = self.tableWidget.selectedRanges()[0]
        for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
            for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                item = self.tableWidget.item(row, col)
                if item and item.text():
                    # every deleted cells info save (row, col, value) 
                    temp_deleted.append((row, col, item.text()))
                    self.tableWidget.takeItem(row, col)
        
        
        if temp_deleted:
            self.deleted_cells.append(temp_deleted)
        
        self.tableWidget.setUpdatesEnabled(True)

    def undo(self):
        
        print("Undo func triggered!")  # Debug

        if self.deleted_cells:
            
            last_deleted = self.deleted_cells.pop()
        
            print(f"Restore Cell Count: {len(last_deleted)}")  # Debug

            for row, col, value in last_deleted:
                self.tableWidget.setItem(row, col, QTableWidgetItem(value))
        else:
            print("No Cells to Restore!")  # Debug 

    def copy(self):
        self.clipboard = []
        selected_range = self.tableWidget.selectedRanges()[0]
        for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
            row_data = []
            for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else '')
            self.clipboard.append(row_data)

    def paste(self):
        if not hasattr(self, 'clipboard') or not self.clipboard:
            return
        
        start_row = self.tableWidget.currentRow()
        start_col = self.tableWidget.currentColumn()

        for row_offset, row_data in enumerate(self.clipboard):
            for col_offset, cell_data in enumerate(row_data):
                self.tableWidget.setItem(start_row + row_offset, start_col + col_offset, QTableWidgetItem(cell_data))

    def cut(self):
        self.copy()
        selected_range = self.tableWidget.selectedRanges()[0]
        for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
            for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                self.tableWidget.takeItem(row, col)

    def importCSV(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if not filePath:
            return

        with open(filePath, 'r', newline='', encoding='utf-8') as file:
            data = list(csv.reader(file))

        rowCount = len(data)
        colCount = max(len(row) for row in data) if data else 0

        self.tableWidget.setRowCount(rowCount)
        self.tableWidget.setColumnCount(colCount)

        for row, rowData in enumerate(data):
            for column, columnData in enumerate(rowData):
                self.tableWidget.setItem(row, column, QTableWidgetItem(columnData))

    def exportCSV(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if not filePath:
            return

        data = []
        for row in range(self.tableWidget.rowCount()):
            row_data = []
            for col in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else '')
            data.append(row_data)

        with open(filePath, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(data)
    
    



    def contextMenuEvent(self, event):
        contextMenu = QMenu(self)
        avgAction = QAction("Average", self)
        avgAction.triggered.connect(self.calculate_average)
        sumAction = QAction("Sum", self)
        sumAction.triggered.connect(self.calculate_sum)
        stdDevAction = QAction("Standard Deviation", self)
        stdDevAction.triggered.connect(self.calculate_std_dev)
        varianceAction = QAction("Variance", self)
        varianceAction.triggered.connect(self.calculate_variance)
        
        contextMenu.addAction(avgAction)
        contextMenu.addAction(sumAction)
        contextMenu.addAction(stdDevAction)
        contextMenu.addAction(varianceAction)
        
        contextMenu.exec_(self.tableWidget.mapToGlobal(event.pos()))

    def get_selected_values(self):
        values = []
        for item in self.tableWidget.selectedItems():
            try:
                val = float(item.text())
                values.append(val)
            except ValueError:
                pass  # Ignore non-numeric values
        return values

    def calculate_average(self):
        values = self.get_selected_values()
        if not values:
            return QMessageBox.information(self, 'Average', 'No valid numbers selected.')
        avg = sum(values) / len(values)
        QMessageBox.information(self, 'Average', f'Average is: {avg:.2f}')

    def calculate_sum(self):
        values = self.get_selected_values()
        if not values:
            return QMessageBox.information(self, 'Sum', 'No valid numbers selected.')
        total = sum(values)
        QMessageBox.information(self, 'Sum', f'Sum is: {total:.2f}')

    def calculate_std_dev(self):
        values = self.get_selected_values()
        if not values:
            return QMessageBox.information(self, 'Standard Deviation', 'No valid numbers selected.')
        mean = sum(values) / len(values)
        variance = sum((x - mean) ** 2 for x in values) / len(values)
        std_dev = variance ** 0.5
        QMessageBox.information(self, 'Standard Deviation', f'Standard Deviation is: {std_dev:.2f}')

    def calculate_variance(self):
        values = self.get_selected_values()
        if not values:
            return QMessageBox.information(self, 'Variance', 'No valid numbers selected.')
        mean = sum(values) / len(values)
        variance = sum((x - mean) ** 2 for x in values) / len(values)
        QMessageBox.information(self, 'Variance', f'Variance is: {variance:.2f}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())


### Easy CSV  v1.0 --- by Berk ÇIKIKCI --- don't forget to follow on Linkedin (Berk ÇIKIKCI) for further projects ----  ###  


