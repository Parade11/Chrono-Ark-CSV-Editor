import sys
import csv
import os
import json
import requests
import subprocess
import platform
import time
import chardet
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTreeWidget, QTreeWidgetItem, QTableWidget,
                             QTableWidgetItem, QSplitter, QMenuBar, QMenu, QFileDialog,
                             QMessageBox, QStatusBar, QInputDialog, QDialog, QLabel,
                             QComboBox, QLineEdit, QPushButton, QFormLayout, QProgressDialog,
                             QTextEdit, QGroupBox)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt6.QtGui import QAction, QKeySequence

DEEPLX_AVAILABLE = True  # We'll use direct API calls


class CSVEditorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.csv_data = []
        self.imported_files = []  # Track imported files
        self.deeplx_process = None  # Track DeepLX process
        self.deeplx_path = Path.home() / ".csv_editor" / "deeplx"
        self.init_ui()
        
        # Auto-start DeepLX if it's installed but not running
        QTimer.singleShot(1000, self.auto_start_deeplx)
        
    def init_ui(self):
        self.setWindowTitle("CSV Editor - Translation Tool")
        self.setGeometry(100, 100, 1200, 700)
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QHBoxLayout(main_widget)
        
        # Create splitter for left and right panels
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left sidebar - File tree
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabel("Imported CSV Files")
        self.file_tree.setMaximumWidth(300)
        self.file_tree.itemClicked.connect(self.on_file_selected)
        self.file_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        splitter.addWidget(self.file_tree)
        
        # Right panel - CSV table
        self.csv_table = QTableWidget()
        self.csv_table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | 
                                       QTableWidget.EditTrigger.EditKeyPressed)
        
        # Enable smooth scrolling
        self.csv_table.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        self.csv_table.setVerticalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        
        # Disable auto column resizing for better horizontal scroll performance
        self.csv_table.horizontalHeader().setStretchLastSection(False)
        
        # Enable multi-selection
        self.csv_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        
        # Enable context menu on column headers
        self.csv_table.horizontalHeader().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.csv_table.horizontalHeader().customContextMenuRequested.connect(self.show_column_context_menu)
        
        # Enable context menu on cells
        self.csv_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.csv_table.customContextMenuRequested.connect(self.show_cell_context_menu)
        
        # Connect cell change signal to track edits
        self.csv_table.itemChanged.connect(self.on_cell_changed)
        
        splitter.addWidget(self.csv_table)
        
        # Set splitter sizes
        splitter.setSizes([250, 950])
        
        layout.addWidget(splitter)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        
    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        import_action = QAction("Import CSV...", self)
        import_action.setShortcut(QKeySequence.StandardKey.Open)
        import_action.triggered.connect(self.import_file)
        file_menu.addAction(import_action)
        
        save_action = QAction("Save", self)
        save_action.setShortcut(QKeySequence.StandardKey.Save)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("Save As...", self)
        save_as_action.setShortcut(QKeySequence.StandardKey.SaveAs)
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        remove_action = QAction("Remove from List", self)
        remove_action.setShortcut("Delete")
        remove_action.triggered.connect(self.remove_selected_file)
        file_menu.addAction(remove_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("Exit", self)
        exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Edit menu
        edit_menu = menubar.addMenu("Edit")
        
        add_column_action = QAction("Add Column...", self)
        add_column_action.setShortcut("Ctrl+Shift+C")
        add_column_action.triggered.connect(self.add_column)
        edit_menu.addAction(add_column_action)
        
        insert_column_action = QAction("Insert Column Before...", self)
        insert_column_action.triggered.connect(self.insert_column_before)
        edit_menu.addAction(insert_column_action)
        
        delete_column_action = QAction("Delete Column", self)
        delete_column_action.setShortcut("Ctrl+Shift+D")
        delete_column_action.triggered.connect(self.delete_column)
        edit_menu.addAction(delete_column_action)
        
        # Settings menu
        settings_menu = menubar.addMenu("Settings")
        
        deeplx_settings_action = QAction("DeepLX Translation Settings...", self)
        deeplx_settings_action.triggered.connect(self.show_deeplx_settings)
        settings_menu.addAction(deeplx_settings_action)
        
    def import_file(self):
        """Import CSV file(s) into the application"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Import CSV Files", "", "CSV Files (*.csv);;All Files (*)"
        )
        
        for file_path in file_paths:
            if file_path and file_path not in self.imported_files:
                self.imported_files.append(file_path)
                self.add_file_to_tree(file_path)
        
        if file_paths:
            self.status_bar.showMessage(f"Imported {len(file_paths)} file(s)")
    
    def add_file_to_tree(self, file_path):
        """Add a file to the tree view"""
        file_name = Path(file_path).name
        item = QTreeWidgetItem([file_name])
        item.setData(0, Qt.ItemDataRole.UserRole, file_path)
        self.file_tree.addTopLevelItem(item)
    
    def remove_selected_file(self):
        """Remove selected file from the list"""
        current_item = self.file_tree.currentItem()
        if current_item:
            file_path = current_item.data(0, Qt.ItemDataRole.UserRole)
            if file_path in self.imported_files:
                self.imported_files.remove(file_path)
            
            index = self.file_tree.indexOfTopLevelItem(current_item)
            self.file_tree.takeTopLevelItem(index)
            
            # Clear table if this was the current file
            if file_path == self.current_file:
                self.csv_table.clear()
                self.csv_table.setRowCount(0)
                self.csv_table.setColumnCount(0)
                self.current_file = None
            
            self.status_bar.showMessage("File removed from list")
        
    def on_file_selected(self, item, column):
        """Handle file selection from tree"""
        file_path = item.data(0, Qt.ItemDataRole.UserRole)
        self.load_csv_file(file_path)
        
    def load_csv_file(self, file_path):
        """Load CSV file into table with encoding and dialect detection"""
        try:
            # Step 1: Detect file encoding
            encoding = 'utf-8'  # Default
            try:
                with open(file_path, 'rb') as f:
                    raw_data = f.read()
                    result = chardet.detect(raw_data)
                    if result['encoding']:
                        encoding = result['encoding']
                        print(f"Detected encoding: {encoding} (confidence: {result['confidence']})")
            except Exception as e:
                print(f"Encoding detection failed, using UTF-8: {e}")
            
            # Step 2: Read sample to detect CSV dialect
            dialect = None
            try:
                with open(file_path, 'r', encoding=encoding, newline='') as f:
                    sample = f.read(8192)  # Read first 8KB
                    sniffer = csv.Sniffer()
                    dialect = sniffer.sniff(sample)
                    print(f"Detected delimiter: '{dialect.delimiter}'")
            except Exception as e:
                print(f"Dialect detection failed, using default: {e}")
            
            # Step 3: Load CSV data with detected settings
            self.csv_data = []
            encodings_to_try = [encoding, 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
            
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc, newline='') as f:
                        if dialect:
                            reader = csv.reader(f, dialect=dialect)
                        else:
                            reader = csv.reader(f)
                        
                        self.csv_data = []
                        for row in reader:
                            self.csv_data.append(row)
                    
                    # Successfully loaded
                    print(f"Successfully loaded with encoding: {enc}")
                    break
                    
                except UnicodeDecodeError:
                    if enc == encodings_to_try[-1]:
                        raise  # Re-raise if last encoding fails
                    continue
                except Exception as e:
                    if enc == encodings_to_try[-1]:
                        raise
                    continue
            
            if not self.csv_data:
                QMessageBox.warning(self, "Empty File", "The CSV file is empty.")
                return
            
            self.current_file = file_path
            self.display_csv_data()
            self.status_bar.showMessage(f"Loaded: {Path(file_path).name} ({encoding})")
            
        except UnicodeDecodeError as e:
            QMessageBox.critical(
                self, "Encoding Error",
                f"Failed to decode file:\n{str(e)}\n\n"
                "The file may be using an unsupported encoding."
            )
        except csv.Error as e:
            QMessageBox.critical(
                self, "CSV Error",
                f"Failed to parse CSV file:\n{str(e)}\n\n"
                "The file may be malformed or not a valid CSV."
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Error",
                f"Failed to load file:\n{str(e)}"
            )
            
    def display_csv_data(self):
        """Display CSV data in table widget"""
        try:
            if not self.csv_data or len(self.csv_data) == 0:
                return
            
            # Get header row and expected column count
            headers = self.csv_data[0]
            expected_col_count = len(headers)
            
            if expected_col_count == 0:
                QMessageBox.warning(self, "Invalid CSV", "CSV file has no columns.")
                return
            
            # Validate and normalize all data rows
            normalized_data = [headers]
            for row_idx, row_data in enumerate(self.csv_data[1:], start=1):
                # Ensure row_data is a list
                if not isinstance(row_data, list):
                    row_data = [str(row_data)]
                
                # Ensure row has correct number of columns
                if len(row_data) < expected_col_count:
                    # Pad with empty strings
                    row_data = list(row_data) + [""] * (expected_col_count - len(row_data))
                elif len(row_data) > expected_col_count:
                    # Truncate extra columns
                    row_data = row_data[:expected_col_count]
                
                normalized_data.append(row_data)
            
            # Update csv_data with normalized data
            self.csv_data = normalized_data
            
            # Set table dimensions (exclude header row from row count)
            self.csv_table.setRowCount(len(self.csv_data) - 1)
            self.csv_table.setColumnCount(expected_col_count)
            
            # Set headers (first row)
            self.csv_table.setHorizontalHeaderLabels(headers)
            
            # Temporarily disconnect itemChanged signal to avoid triggering during load
            self.csv_table.itemChanged.disconnect(self.on_cell_changed)
            
            # Populate table (skip header row)
            for row_idx, row_data in enumerate(self.csv_data[1:]):
                # Double-check bounds
                if row_idx >= self.csv_table.rowCount():
                    break
                
                for col_idx in range(min(len(row_data), expected_col_count)):
                    try:
                        # Safely get cell data with bounds checking
                        cell_data = row_data[col_idx] if col_idx < len(row_data) else ""
                        item = QTableWidgetItem(str(cell_data) if cell_data is not None else "")
                        self.csv_table.setItem(row_idx, col_idx, item)
                    except (IndexError, ValueError) as e:
                        # If there's an error with a specific cell, just use empty string
                        print(f"Warning: Error setting cell [{row_idx}, {col_idx}]: {e}")
                        self.csv_table.setItem(row_idx, col_idx, QTableWidgetItem(""))
            
            # Reconnect itemChanged signal after loading
            self.csv_table.itemChanged.connect(self.on_cell_changed)
        
        except Exception as e:
            QMessageBox.critical(
                self, "Display Error",
                f"Failed to display CSV data:\n{str(e)}\n\nThe file may be malformed."
            )
            return
        
        # Adjust column widths with minimum width for better scrolling
        self.csv_table.resizeColumnsToContents()
        
        # Set minimum column width for smoother scrolling
        for col in range(self.csv_table.columnCount()):
            current_width = self.csv_table.columnWidth(col)
            self.csv_table.setColumnWidth(col, max(current_width, 100))
        

            
    def on_cell_changed(self, item):
        """Handle individual cell changes and update csv_data"""
        if not self.csv_data or not item:
            return
        
        try:
            row = item.row()
            col = item.column()
            
            # Update csv_data (row + 1 because csv_data[0] is headers)
            data_row = row + 1
            
            # Ensure csv_data has enough rows
            while len(self.csv_data) <= data_row:
                # Add empty row with correct number of columns
                self.csv_data.append([""] * self.csv_table.columnCount())
            
            # Ensure the row has enough columns
            while len(self.csv_data[data_row]) <= col:
                self.csv_data[data_row].append("")
            
            # Update the specific cell
            self.csv_data[data_row][col] = item.text()
            
        except Exception as e:
            print(f"Error updating cell data: {e}")
    
    def sync_csv_data_from_table(self):
        """Synchronize self.csv_data with current table contents"""
        if not self.csv_table.rowCount() or not self.csv_table.columnCount():
            return
        
        # Collect headers from table
        headers = []
        for col in range(self.csv_table.columnCount()):
            header_item = self.csv_table.horizontalHeaderItem(col)
            headers.append(header_item.text() if header_item else f"Column {col}")
        
        # Collect all rows
        data = [headers]
        for row in range(self.csv_table.rowCount()):
            row_data = []
            for col in range(self.csv_table.columnCount()):
                item = self.csv_table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        
        # Update csv_data
        self.csv_data = data
    
    def save_file(self):
        """Save current CSV file"""
        if not self.current_file:
            self.save_file_as()
            return
        
        try:
            # Sync data from table to csv_data
            self.sync_csv_data_from_table()
            
            # Write to file with quotes around all non-empty fields
            with open(self.current_file, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC, lineterminator='\n')
                writer.writerows(self.csv_data)
            
            self.status_bar.showMessage(f"Saved: {Path(self.current_file).name}")
            QMessageBox.information(self, "Success", "File saved successfully!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file:\n{str(e)}")
            
    def save_file_as(self):
        """Save CSV file with new name"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save CSV File As", "", "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            old_file = self.current_file
            self.current_file = file_path
            self.save_file()
            
            # Update imported files list
            if old_file in self.imported_files:
                self.imported_files.remove(old_file)
            if file_path not in self.imported_files:
                self.imported_files.append(file_path)
            
            # Refresh tree
            self.file_tree.clear()
            for imported_file in self.imported_files:
                self.add_file_to_tree(imported_file)
    
    def add_column(self):
        """Add a new column at the end"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        column_name, ok = QInputDialog.getText(
            self, "Add Column", "Enter column name (e.g., English):"
        )
        
        if ok and column_name:
            # Add column to table
            col_count = self.csv_table.columnCount()
            self.csv_table.insertColumn(col_count)
            self.csv_table.setHorizontalHeaderItem(col_count, QTableWidgetItem(column_name))
            
            # Sync data to ensure consistency
            self.sync_csv_data_from_table()
            
            # Set column width
            self.csv_table.setColumnWidth(col_count, 150)
            
            self.status_bar.showMessage(f"Added column: {column_name}")
    
    def insert_column_before(self):
        """Insert a new column before the selected column"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        current_col = self.csv_table.currentColumn()
        if current_col < 0:
            QMessageBox.warning(self, "No Selection", "Please select a column first.")
            return
        
        column_name, ok = QInputDialog.getText(
            self, "Insert Column", "Enter column name:"
        )
        
        if ok and column_name:
            # Insert column in table
            self.csv_table.insertColumn(current_col)
            self.csv_table.setHorizontalHeaderItem(current_col, QTableWidgetItem(column_name))
            
            # Sync data to ensure consistency
            self.sync_csv_data_from_table()
            
            # Set column width
            self.csv_table.setColumnWidth(current_col, 150)
            
            self.status_bar.showMessage(f"Inserted column: {column_name}")
    
    def delete_column(self):
        """Delete the selected column"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        current_col = self.csv_table.currentColumn()
        if current_col < 0:
            QMessageBox.warning(self, "No Selection", "Please select a column to delete.")
            return
        
        # Safely get column name with null check
        header_item = self.csv_table.horizontalHeaderItem(current_col)
        column_name = header_item.text() if header_item else f"Column {current_col}"
        
        reply = QMessageBox.question(
            self, "Delete Column",
            f"Are you sure you want to delete column '{column_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Remove column from table
            self.csv_table.removeColumn(current_col)
            
            # Sync data to ensure consistency
            self.sync_csv_data_from_table()
            
            self.status_bar.showMessage(f"Deleted column: {column_name}")
    
    def show_deeplx_settings(self):
        """Show DeepLX settings dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("DeepLX Translation Settings")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(400)
        
        layout = QVBoxLayout()
        
        # Status group
        status_group = QGroupBox("DeepLX Status")
        status_layout = QVBoxLayout()
        
        status_label = QLabel("Checking DeepLX status...")
        status_layout.addWidget(status_label)
        
        status_group.setLayout(status_layout)
        layout.addWidget(status_group)
        
        # Check if DeepLX is running
        def check_deeplx_status():
            try:
                response = requests.get("http://127.0.0.1:1188", timeout=2)
                status_label.setText("✓ DeepLX is running on port 1188")
                status_label.setStyleSheet("color: green; font-weight: bold;")
                start_btn.setEnabled(False)
                stop_btn.setEnabled(True)
                return True
            except:
                status_label.setText("✗ DeepLX is not running")
                status_label.setStyleSheet("color: red; font-weight: bold;")
                start_btn.setEnabled(True)
                stop_btn.setEnabled(False)
                return False
        
        # Control buttons
        button_group = QGroupBox("DeepLX Control")
        button_layout = QVBoxLayout()
        
        # Download and setup button
        setup_btn = QPushButton("Download && Setup DeepLX")
        setup_btn.clicked.connect(lambda: self.download_deeplx(dialog, status_label, check_deeplx_status))
        button_layout.addWidget(setup_btn)
        
        # Start button
        start_btn = QPushButton("Start DeepLX Server")
        start_btn.clicked.connect(lambda: self.start_deeplx(status_label, check_deeplx_status))
        button_layout.addWidget(start_btn)
        
        # Stop button
        stop_btn = QPushButton("Stop DeepLX Server")
        stop_btn.clicked.connect(lambda: self.stop_deeplx(status_label, check_deeplx_status))
        button_layout.addWidget(stop_btn)
        
        button_group.setLayout(button_layout)
        layout.addWidget(button_group)
        
        # Info text
        info_group = QGroupBox("Information")
        info_layout = QVBoxLayout()
        
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        info_text.setMaximumHeight(150)
        info_text.setPlainText(
            "1. Click 'Download & Setup DeepLX' to automatically download it\n"
            "2. Click 'Start DeepLX Server' to run it locally\n"
            "3. Use the translation feature in the app\n\n"
            "The server runs on http://127.0.0.1:1188"
        )
        info_layout.addWidget(info_text)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        
        dialog.setLayout(layout)
        
        # Initial status check
        check_deeplx_status()
        
        dialog.exec()
    
    def download_deeplx(self, parent_dialog, status_label, check_status_func):
        """Download DeepLX binary"""
        system = platform.system()
        machine = platform.machine().lower()
        
        # Determine download URL based on OS
        # First, get the latest release info
        try:
            release_response = requests.get(
                "https://api.github.com/repos/OwO-Network/DeepLX/releases/latest",
                timeout=10
            )
            release_response.raise_for_status()
            release_data = release_response.json()
            assets = release_data.get("assets", [])
            
            # Find the correct asset for this platform
            if system == "Windows":
                if "amd64" in machine or "x86_64" in machine:
                    asset_name = "deeplx_windows_amd64.exe"
                elif "386" in machine or "x86" in machine:
                    asset_name = "deeplx_windows_386.exe"
                else:
                    QMessageBox.warning(parent_dialog, "Unsupported", "Your Windows architecture is not supported.")
                    return
                exe_name = "deeplx.exe"
            elif system == "Linux":
                if "amd64" in machine or "x86_64" in machine:
                    asset_name = "deeplx_linux_amd64"
                elif "arm64" in machine or "aarch64" in machine:
                    asset_name = "deeplx_linux_arm64"
                elif "386" in machine or "i686" in machine:
                    asset_name = "deeplx_linux_386"
                else:
                    QMessageBox.warning(parent_dialog, "Unsupported", "Your Linux architecture is not supported.")
                    return
                exe_name = "deeplx"
            elif system == "Darwin":  # macOS
                if "arm" in machine or "aarch64" in machine:
                    asset_name = "deeplx_darwin_arm64"
                else:
                    asset_name = "deeplx_darwin_amd64"
                exe_name = "deeplx"
            else:
                QMessageBox.warning(parent_dialog, "Unsupported", f"Your operating system ({system}) is not supported.")
                return
            
            # Find the download URL for the asset
            url = None
            for asset in assets:
                if asset.get("name") == asset_name:
                    url = asset.get("browser_download_url")
                    break
            
            if not url:
                QMessageBox.warning(
                    parent_dialog, "Not Found",
                    f"Could not find {asset_name} in the latest release.\n\n"
                    "Please download manually from:\n"
                    "https://github.com/OwO-Network/DeepLX/releases"
                )
                return
                
        except Exception as e:
            QMessageBox.critical(
                parent_dialog, "Error",
                f"Failed to get release information:\n{str(e)}\n\n"
                "Please download manually from:\n"
                "https://github.com/OwO-Network/DeepLX/releases"
            )
            return
        
        # Create progress dialog
        progress = QProgressDialog("Downloading DeepLX...", "Cancel", 0, 100, parent_dialog)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        
        try:
            # Create directory
            self.deeplx_path.mkdir(parents=True, exist_ok=True)
            
            # Download file directly (it's an executable, not a zip)
            status_label.setText("Downloading DeepLX...")
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            exe_path = self.deeplx_path / exe_name
            
            downloaded = 0
            with open(exe_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if progress.wasCanceled():
                        return
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        progress.setValue(int(downloaded * 100 / total_size))
            
            # Make executable on Unix systems
            if system != "Windows":
                os.chmod(exe_path, 0o755)
            
            progress.setValue(100)
            status_label.setText("✓ DeepLX downloaded successfully!")
            status_label.setStyleSheet("color: green; font-weight: bold;")
            
            QMessageBox.information(
                parent_dialog, "Success",
                "DeepLX has been downloaded successfully!\n\n"
                "Click 'Start DeepLX Server' to run it."
            )
            
        except Exception as e:
            status_label.setText(f"✗ Download failed: {str(e)}")
            status_label.setStyleSheet("color: red; font-weight: bold;")
            QMessageBox.critical(parent_dialog, "Download Failed", f"Failed to download DeepLX:\n{str(e)}")
    
    def start_deeplx(self, status_label, check_status_func):
        """Start DeepLX server"""
        system = platform.system()
        exe_name = "deeplx.exe" if system == "Windows" else "deeplx"
        exe_path = self.deeplx_path / exe_name
        
        if not exe_path.exists():
            QMessageBox.warning(
                self, "DeepLX Not Found",
                "DeepLX is not installed.\n\n"
                "Click 'Download & Setup DeepLX' first."
            )
            return
        
        try:
            # Start DeepLX process
            if system == "Windows":
                self.deeplx_process = subprocess.Popen(
                    [str(exe_path)],
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                self.deeplx_process = subprocess.Popen(
                    [str(exe_path)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            
            status_label.setText("Starting DeepLX...")
            status_label.setStyleSheet("color: orange; font-weight: bold;")
            
            # Wait a moment and check status
            QTimer.singleShot(2000, check_status_func)
            
            self.status_bar.showMessage("DeepLX server started")
            
        except Exception as e:
            status_label.setText(f"✗ Failed to start: {str(e)}")
            status_label.setStyleSheet("color: red; font-weight: bold;")
            QMessageBox.critical(self, "Start Failed", f"Failed to start DeepLX:\n{str(e)}")
    
    def stop_deeplx(self, status_label, check_status_func):
        """Stop DeepLX server"""
        if self.deeplx_process:
            try:
                self.deeplx_process.terminate()
                self.deeplx_process.wait(timeout=5)
                self.deeplx_process = None
                status_label.setText("✗ DeepLX stopped")
                status_label.setStyleSheet("color: red; font-weight: bold;")
                self.status_bar.showMessage("DeepLX server stopped")
                check_status_func()
            except Exception as e:
                QMessageBox.warning(self, "Stop Failed", f"Failed to stop DeepLX:\n{str(e)}")
        else:
            QMessageBox.information(self, "Not Running", "DeepLX is not running from this application.")
    
    def auto_start_deeplx(self):
        """Automatically start DeepLX if installed but not running"""
        # Check if DeepLX is already running
        try:
            response = requests.get("http://127.0.0.1:1188", timeout=1)
            # Already running, no need to start
            return
        except:
            pass
        
        # Check if DeepLX is installed
        system = platform.system()
        exe_name = "deeplx.exe" if system == "Windows" else "deeplx"
        exe_path = self.deeplx_path / exe_name
        
        if exe_path.exists():
            # Start it automatically
            try:
                if system == "Windows":
                    self.deeplx_process = subprocess.Popen(
                        [str(exe_path)],
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    self.deeplx_process = subprocess.Popen(
                        [str(exe_path)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                self.status_bar.showMessage("DeepLX server started automatically")
            except Exception as e:
                print(f"Failed to auto-start DeepLX: {e}")
    
    def closeEvent(self, event):
        """Clean up when closing the application"""
        if self.deeplx_process:
            try:
                self.deeplx_process.terminate()
                self.deeplx_process.wait(timeout=5)
            except:
                pass
        event.accept()
    
    def translate_selected_cells(self, selected_items, source_lang, target_lang):
        """Translate selected cells"""
        total_cells = len(selected_items)
        
        # DeepLX endpoint
        endpoint = "http://127.0.0.1:1188/translate"
        
        # Create progress dialog
        progress = QProgressDialog("Translating selected cells...", "Cancel", 0, total_cells, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        
        translated_count = 0
        failed_count = 0
        
        for idx, item in enumerate(selected_items):
            if progress.wasCanceled():
                break
            
            progress.setValue(idx)
            progress.setLabelText(f"Translating cell {idx + 1} of {total_cells}...")
            QApplication.processEvents()
            
            # Get cell text
            cell_text = item.text().strip()
            if not cell_text:
                continue
            
            # Translate
            try:
                data = {
                    "text": cell_text,
                    "source_lang": source_lang,
                    "target_lang": target_lang
                }
                
                response = requests.post(
                    endpoint,
                    json=data,
                    headers={"Content-Type": "application/json"},
                    timeout=15
                )
                
                if response.status_code == 200:
                    result = response.json()
                    translated_text = result.get("data") or result.get("text", "")
                    
                    if translated_text and not translated_text.startswith("http"):
                        item.setText(translated_text)
                        translated_count += 1
                        # Add delay to avoid rate limiting
                        time.sleep(0.5)
                    else:
                        failed_count += 1
                else:
                    failed_count += 1
                    print(f"HTTP {response.status_code}: {response.text}")
                    
            except Exception as e:
                failed_count += 1
                print(f"Translation failed: {e}")
        
        progress.setValue(total_cells)
        
        # Show summary
        if translated_count > 0:
            message = f"Translation complete!\n\nTranslated: {translated_count} cells"
            if failed_count > 0:
                message += f"\nFailed: {failed_count} cells"
            QMessageBox.information(self, "Translation Complete", message)
            self.status_bar.showMessage(f"Translated {translated_count} cells")
        else:
            QMessageBox.warning(
                self, "Translation Failed",
                "No cells were translated.\n\n"
                "Make sure DeepLX is running (Settings > DeepLX Translation Settings)"
            )
    
    def show_column_context_menu(self, position):
        """Show context menu for column header"""
        if not self.current_file:
            return
        
        column = self.csv_table.horizontalHeader().logicalIndexAt(position)
        if column < 0:
            return
        
        menu = QMenu()
        translate_action = QAction("Translate Entire Column with DeepL...", self)
        translate_action.triggered.connect(lambda: self.show_translate_dialog(column))
        menu.addAction(translate_action)
        
        menu.exec(self.csv_table.horizontalHeader().mapToGlobal(position))
    
    def show_cell_context_menu(self, position):
        """Show context menu for cells"""
        if not self.current_file:
            return
        
        selected_items = self.csv_table.selectedItems()
        if not selected_items:
            return
        
        menu = QMenu()
        
        # Translate selected cells
        translate_action = QAction(f"Translate Selected Cells ({len(selected_items)})...", self)
        translate_action.triggered.connect(self.show_translate_cells_dialog)
        menu.addAction(translate_action)
        
        menu.exec(self.csv_table.viewport().mapToGlobal(position))
    
    def show_translate_dialog(self, target_column):
        """Show dialog to configure translation"""
        
        # Get target column name
        target_header = self.csv_table.horizontalHeaderItem(target_column)
        target_name = target_header.text() if target_header else f"Column {target_column}"
        
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Translate to: {target_name}")
        dialog.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        # Show target column
        target_label = QLabel(f"<b>{target_name}</b>")
        layout.addRow("Target Column:", target_label)
        
        # Source column selection
        source_combo = QComboBox()
        for col in range(self.csv_table.columnCount()):
            if col != target_column:
                header = self.csv_table.horizontalHeaderItem(col)
                source_combo.addItem(header.text() if header else f"Column {col}", col)
        layout.addRow("Source Column:", source_combo)
        
        # Auto-detect languages from column names
        def detect_lang_code(name):
            name_lower = name.lower()
            if "english" in name_lower or name_lower == "en":
                return "EN"
            elif "chinese" in name_lower or name_lower == "zh":
                return "ZH"
            elif "japanese" in name_lower or name_lower == "ja":
                return "JA"
            elif "korean" in name_lower or name_lower == "ko":
                return "KO"
            return ""
        
        # Source language
        default_source = detect_lang_code(source_combo.currentText()) or "EN"
        source_lang = QLineEdit(default_source)
        source_lang.setPlaceholderText("e.g., EN, ZH, JA, KO")
        layout.addRow("Source Language:", source_lang)
        
        # Target language
        default_target = detect_lang_code(target_name) or "KO"
        target_lang = QLineEdit(default_target)
        target_lang.setPlaceholderText("e.g., EN, ZH, JA, KO")
        layout.addRow("Target Language:", target_lang)
        
        # Update source language when source column changes
        def update_source_lang():
            lang = detect_lang_code(source_combo.currentText())
            if lang:
                source_lang.setText(lang)
        source_combo.currentTextChanged.connect(update_source_lang)
        
        # Info label with status check
        def check_deeplx_for_info():
            try:
                requests.get("http://127.0.0.1:1188", timeout=1)
                return "✓ DeepLX is running - ready to translate!"
            except:
                return "⚠ DeepLX is not running. Go to Settings > DeepLX Translation Settings to start it."
        
        info_label = QLabel(check_deeplx_for_info())
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow(info_label)
        
        # Buttons
        button_layout = QHBoxLayout()
        translate_btn = QPushButton("Translate")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(translate_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow(button_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        translate_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            source_col = source_combo.currentData()
            src_lang = source_lang.text().strip().upper()
            tgt_lang = target_lang.text().strip().upper()
            
            if not src_lang or not tgt_lang:
                QMessageBox.warning(self, "Invalid Input", "Please enter both source and target languages.")
                return
            
            # Start translation
            self.translate_column(source_col, target_column, src_lang, tgt_lang)
    
    def show_translate_cells_dialog(self):
        """Show dialog to translate selected cells"""
        selected_items = self.csv_table.selectedItems()
        if not selected_items:
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Translate Selected Cells")
        dialog.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        # Show selection info
        info_label = QLabel(f"<b>{len(selected_items)} cells selected</b>")
        layout.addRow("Selection:", info_label)
        
        # Source language
        source_lang = QLineEdit("EN")
        source_lang.setPlaceholderText("e.g., EN, ZH, JA, KO")
        layout.addRow("Source Language:", source_lang)
        
        # Target language
        target_lang = QLineEdit("JA")
        target_lang.setPlaceholderText("e.g., EN, ZH, JA, KO")
        layout.addRow("Target Language:", target_lang)
        
        # Info label with status check
        def check_deeplx_for_info():
            try:
                requests.get("http://127.0.0.1:1188", timeout=1)
                return "✓ DeepLX is running - ready to translate!"
            except:
                return "⚠ DeepLX is not running. Go to Settings > DeepLX Translation Settings to start it."
        
        info_label2 = QLabel(check_deeplx_for_info())
        info_label2.setWordWrap(True)
        info_label2.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow(info_label2)
        
        # Buttons
        button_layout = QHBoxLayout()
        translate_btn = QPushButton("Translate")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(translate_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow(button_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        translate_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            src_lang = source_lang.text().strip().upper()
            tgt_lang = target_lang.text().strip().upper()
            
            if not src_lang or not tgt_lang:
                QMessageBox.warning(self, "Invalid Input", "Please enter both source and target languages.")
                return
            
            # Start translation
            self.translate_selected_cells(selected_items, src_lang, tgt_lang)
    
    def translate_column(self, source_col, target_col, source_lang, target_lang):
        """Translate all cells from source column to target column using DeepLX API"""
        row_count = self.csv_table.rowCount()
        
        # DeepLX API endpoints to try (prioritize local)
        endpoints = [
            "http://127.0.0.1:1188/translate",  # Local DeepLX (most reliable)
        ]
        
        # Create progress dialog
        progress = QProgressDialog("Translating...", "Cancel", 0, row_count, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        
        translated_count = 0
        failed_count = 0
        current_endpoint = 0
        
        for row in range(row_count):
            if progress.wasCanceled():
                break
            
            progress.setValue(row)
            progress.setLabelText(f"Translating row {row + 1} of {row_count}...")
            QApplication.processEvents()
            
            # Get source text
            source_item = self.csv_table.item(row, source_col)
            if not source_item or not source_item.text().strip():
                continue
            
            source_text = source_item.text()
            
            # Try translation with fallback endpoints
            translated = False
            for attempt in range(len(endpoints)):
                endpoint = endpoints[(current_endpoint + attempt) % len(endpoints)]
                
                try:
                    # Prepare request data
                    data = {
                        "text": source_text,
                        "source_lang": source_lang,
                        "target_lang": target_lang
                    }
                    
                    # Make request to DeepLX API
                    response = requests.post(
                        endpoint,
                        json=data,
                        headers={"Content-Type": "application/json"},
                        timeout=15
                    )
                    
                    if response.status_code == 200:
                        result = response.json()
                        
                        # Extract translated text
                        translated_text = result.get("data") or result.get("text", "")
                        
                        # Validate translation (check if it's not a URL or error message)
                        if translated_text and not translated_text.startswith("http") and len(translated_text) > 0:
                            # Update target cell
                            target_item = self.csv_table.item(row, target_col)
                            if target_item:
                                target_item.setText(translated_text)
                            else:
                                self.csv_table.setItem(row, target_col, QTableWidgetItem(translated_text))
                            
                            translated_count += 1
                            translated = True
                            break
                        else:
                            # Invalid response, try next endpoint
                            print(f"Invalid translation from {endpoint}: {translated_text}")
                            print(f"Full response: {result}")
                            continue
                    elif response.status_code == 429:
                        # Rate limited, try next endpoint
                        print(f"Rate limited by {endpoint}")
                        current_endpoint = (current_endpoint + 1) % len(endpoints)
                        continue
                    else:
                        print(f"HTTP {response.status_code} from {endpoint}: {response.text}")
                    
                except requests.exceptions.RequestException as e:
                    print(f"Endpoint {endpoint} failed: {e}")
                    continue
                except Exception as e:
                    print(f"Error with endpoint {endpoint}: {e}")
                    continue
            
            if not translated:
                failed_count += 1
            else:
                # Add a small delay between successful translations to avoid rate limiting
                time.sleep(0.5)  # 500ms delay
        
        progress.setValue(row_count)
        
        # Show summary
        if translated_count > 0:
            message = f"Translation complete!\n\nTranslated: {translated_count} rows"
            if failed_count > 0:
                message += f"\nFailed: {failed_count} rows"
                if "503" in str(failed_count):  # Rate limit hint
                    message += "\n\nNote: Some translations failed due to rate limiting."
            QMessageBox.information(self, "Translation Complete", message)
            self.status_bar.showMessage(f"Translated {translated_count} rows")
        else:
            # Check if it's a rate limit issue
            try:
                test_response = requests.get("http://127.0.0.1:1188", timeout=2)
                if test_response.status_code == 503 or "blocked" in test_response.text.lower():
                    QMessageBox.warning(
                        self, "Rate Limited", 
                        "Your IP has been temporarily blocked by DeepL due to too many requests.\n\n"
                        "Please wait a few minutes before trying again.\n\n"
                        "The app now adds a 500ms delay between translations to prevent this."
                    )
                    return
            except:
                pass
            
            QMessageBox.warning(
                self, "Translation Failed", 
                "No rows were translated.\n\n"
                "DeepLX server is not running or not responding.\n\n"
                "Please start DeepLX:\n"
                "1. Go to Settings > DeepLX Translation Settings\n"
                "2. Click 'Download & Setup DeepLX' (if not already done)\n"
                "3. Click 'Start DeepLX Server'\n"
                "4. Try translation again"
            )



def main():
    app = QApplication(sys.argv)
    window = CSVEditorWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
