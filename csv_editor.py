import sys
import csv
import os
import json
import requests
import subprocess
import platform
import time
import chardet
import re
import random
from pathlib import Path
from enum import Enum
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTreeWidget, QTreeWidgetItem, QTableWidget,
                             QTableWidgetItem, QSplitter, QMenuBar, QMenu, QFileDialog,
                             QMessageBox, QStatusBar, QInputDialog, QDialog, QLabel,
                             QComboBox, QLineEdit, QPushButton, QFormLayout, QProgressDialog,
                             QTextEdit, QGroupBox, QCheckBox, QSpinBox, QDoubleSpinBox,
                             QListWidget, QListWidgetItem, QHBoxLayout)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal, QThreadPool, QRunnable
from PyQt6.QtGui import QAction, QKeySequence, QColor

# New imports for translation services
try:
    from deep_translator import GoogleTranslator, MyMemoryTranslator
    DEEP_TRANSLATOR_AVAILABLE = True
except ImportError:
    DEEP_TRANSLATOR_AVAILABLE = False

try:
    from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
    TENACITY_AVAILABLE = True
except ImportError:
    TENACITY_AVAILABLE = False

class TranslationService(Enum):
    GOOGLE = "google"
    MYMEMORY = "mymemory"


class CSVEditorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.csv_data = []
        self.imported_files = []  # Track imported files
        self.search_results = []  # Store search results
        self.current_search_index = -1  # Current position in search results
        self.search_active = False  # Flag to prevent clearing during navigation
        self.file_data_cache = {}  # Cache data for all imported files {file_path: csv_data}
        self.modified_files = set()  # Track which files have been modified
        self.config = {}
        self.translation_log = []
        self.endpoint_status = {}
        self.circuit_breaker = {}  # Track failures per endpoint
        self.last_successful_translation = None
        self.load_config()
        self.init_ui()
        
    def load_config(self):
        config_path = Path.home() / ".csv_editor" / "config.json"
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    self.config = json.load(f)
            except:
                self.config = {}
        else:
            self.config = {}
        
        # Set defaults
        self.config.setdefault('preferred_service', 'google')
        self.config.setdefault('request_timeout', 15)
        self.config.setdefault('retry_count', 3)
        self.config.setdefault('base_delay', 8.0)
        self.config.setdefault('enabled_services', ['google', 'mymemory'])
        self.config.setdefault('priority_order', ['google', 'mymemory'])
        self.config.setdefault('last_successful_endpoints', {})
        self.config.setdefault('circuit_breaker_threshold', 5)
        self.config.setdefault('circuit_breaker_timeout', 300)  # 5 minutes
        
        # Clean up invalid services from config
        valid_services = [s.value for s in TranslationService]
        self.config['enabled_services'] = [s for s in self.config['enabled_services'] if s in valid_services]
        self.config['priority_order'] = [s for s in self.config['priority_order'] if s in valid_services]
        
        # Ensure at least one service is enabled
        if not self.config['enabled_services']:
            self.config['enabled_services'] = ['google']
        if not self.config['priority_order']:
            self.config['priority_order'] = ['google']
    
    def save_config(self):
        config_path = Path.home() / ".csv_editor" / "config.json"
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w') as f:
            json.dump(self.config, f, indent=2)
    
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
        
        # Enable column reordering via drag-and-drop
        self.csv_table.horizontalHeader().setSectionsMovable(True)
        self.csv_table.horizontalHeader().setDragEnabled(True)
        self.csv_table.horizontalHeader().setDragDropMode(self.csv_table.horizontalHeader().DragDropMode.InternalMove)
        
        # Connect column moved signal
        self.csv_table.horizontalHeader().sectionMoved.connect(self.on_column_moved)
        
        # Enable context menu on column headers
        self.csv_table.horizontalHeader().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.csv_table.horizontalHeader().customContextMenuRequested.connect(self.show_column_context_menu)
        
        # Enable context menu on cells
        self.csv_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.csv_table.customContextMenuRequested.connect(self.show_cell_context_menu)
        
        # Connect cell change signal to track edits
        self.csv_table.itemChanged.connect(self.on_cell_changed)
        
        # Connect cell click to clear search highlights
        self.csv_table.cellClicked.connect(self.on_cell_clicked)
        
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
        
        save_all_action = QAction("Save All", self)
        save_all_action.setShortcut("Ctrl+Alt+S")
        save_all_action.triggered.connect(self.save_all_files)
        file_menu.addAction(save_all_action)
        
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
        
        edit_menu.addSeparator()
        
        # Row operations
        add_row_action = QAction("Add Row", self)
        add_row_action.setShortcut("Ctrl+Shift+R")
        add_row_action.triggered.connect(self.add_row)
        edit_menu.addAction(add_row_action)
        
        insert_row_action = QAction("Insert Row Before...", self)
        insert_row_action.setShortcut("Ctrl+Shift+I")
        insert_row_action.triggered.connect(self.insert_row_before)
        edit_menu.addAction(insert_row_action)
        
        delete_row_action = QAction("Delete Row(s)", self)
        delete_row_action.setShortcut("Ctrl+Shift+Delete")
        delete_row_action.triggered.connect(self.delete_row)
        edit_menu.addAction(delete_row_action)
        
        edit_menu.addSeparator()
        
        # Validation
        validate_action = QAction("Validate Data...", self)
        validate_action.triggered.connect(self.show_validation_dialog)
        edit_menu.addAction(validate_action)
        
        # Find menu
        find_menu = menubar.addMenu("Find")
        
        find_action = QAction("Find...", self)
        find_action.setShortcut(QKeySequence.StandardKey.Find)
        find_action.triggered.connect(self.show_find_dialog)
        find_menu.addAction(find_action)
        
        find_next_action = QAction("Find Next", self)
        find_next_action.setShortcut("F3")
        find_next_action.triggered.connect(self.find_next)
        find_menu.addAction(find_next_action)
        
        find_prev_action = QAction("Find Previous", self)
        find_prev_action.setShortcut("Shift+F3")
        find_prev_action.triggered.connect(self.find_previous)
        find_menu.addAction(find_prev_action)
        
        find_menu.addSeparator()
        
        replace_action = QAction("Replace...", self)
        replace_action.setShortcut(QKeySequence.StandardKey.Replace)
        replace_action.triggered.connect(self.show_replace_dialog)
        find_menu.addAction(replace_action)
        
        # Settings menu
        settings_menu = menubar.addMenu("Settings")
        
        translation_config_action = QAction("Translation Services Configuration...", self)
        translation_config_action.triggered.connect(self.show_translation_config_dialog)
        settings_menu.addAction(translation_config_action)
        
        view_log_action = QAction("View Translation Log...", self)
        view_log_action.triggered.connect(self.show_translation_log)
        settings_menu.addAction(view_log_action)
        
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
        # Save current file data to cache before switching
        if self.current_file and self.csv_data:
            self.sync_csv_data_from_table()
            self.file_data_cache[self.current_file] = self.csv_data.copy()
        
        file_path = item.data(0, Qt.ItemDataRole.UserRole)
        self.load_csv_file(file_path)
        
    def load_csv_file(self, file_path):
        """Load CSV file into table with encoding and dialect detection"""
        try:
            # Check if we have cached data for this file
            if file_path in self.file_data_cache:
                print(f"Loading from cache: {file_path}")
                self.csv_data = self.file_data_cache[file_path].copy()
                self.current_file = file_path
                self.display_csv_data()
                self.status_bar.showMessage(f"Loaded: {Path(file_path).name} (from cache)")
                return
            
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
            
            # Cache the loaded data
            self.file_data_cache[file_path] = self.csv_data.copy()
            
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
        

            
    def on_column_moved(self, logical_index, old_visual_index, new_visual_index):
        """Handle column reordering and update csv_data"""
        if not self.csv_data:
            return
        
        try:
            print(f"Column moved: logical={logical_index}, from={old_visual_index}, to={new_visual_index}")
            
            # Get the current visual order of columns
            header = self.csv_table.horizontalHeader()
            col_count = self.csv_table.columnCount()
            
            # Create a mapping of visual position to logical index
            visual_to_logical = []
            for visual_pos in range(col_count):
                logical_pos = header.logicalIndex(visual_pos)
                visual_to_logical.append(logical_pos)
            
            print(f"Visual to logical mapping: {visual_to_logical}")
            
            # Reorder csv_data based on the new visual order
            new_csv_data = []
            for row_data in self.csv_data:
                new_row = [row_data[logical_idx] for logical_idx in visual_to_logical]
                new_csv_data.append(new_row)
            
            self.csv_data = new_csv_data
            
            # Mark as modified
            if self.current_file:
                self.modified_files.add(self.current_file)
                self.file_data_cache[self.current_file] = self.csv_data.copy()
                self.update_file_tree_indicators()
            
            self.status_bar.showMessage(f"Column reordered: {self.csv_data[0][new_visual_index]}")
            
        except Exception as e:
            print(f"Error reordering columns: {e}")
            QMessageBox.warning(self, "Reorder Error", f"Failed to reorder columns:\n{str(e)}")
    
    def on_cell_clicked(self, row, col):
        """Handle cell click - clear search highlights if not navigating"""
        # Don't clear if we're actively navigating search results
        if self.search_results and not self.search_active:
            self.clear_search_highlights()
            self.search_results = []
            self.current_search_index = -1
            self.status_bar.showMessage("Search cleared")
    
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
            
            # Mark current file as modified and update cache
            if self.current_file:
                self.modified_files.add(self.current_file)
                self.file_data_cache[self.current_file] = self.csv_data.copy()
                self.update_file_tree_indicators()
            
        except Exception as e:
            print(f"Error updating cell data: {e}")
    
    def validate_csv_data(self):
        """Validate CSV data integrity and return issues"""
        issues = []
        
        if not self.csv_data or len(self.csv_data) == 0:
            issues.append("CSV data is empty")
            return issues
        
        # Check headers
        headers = self.csv_data[0]
        if not headers or len(headers) == 0:
            issues.append("No column headers found")
            return issues
        
        # Check for empty headers
        for idx, header in enumerate(headers):
            if not header or not str(header).strip():
                issues.append(f"Column {idx + 1} has an empty header")
        
        # Check for duplicate headers
        header_counts = {}
        for idx, header in enumerate(headers):
            header_str = str(header).strip()
            if header_str in header_counts:
                issues.append(f"Duplicate header '{header_str}' found in columns {header_counts[header_str] + 1} and {idx + 1}")
            else:
                header_counts[header_str] = idx
        
        # Check column count consistency
        expected_col_count = len(headers)
        for row_idx, row_data in enumerate(self.csv_data[1:], start=1):
            if len(row_data) != expected_col_count:
                issues.append(
                    f"Row {row_idx} has {len(row_data)} columns, expected {expected_col_count}"
                )
        
        # Check for completely empty rows
        for row_idx, row_data in enumerate(self.csv_data[1:], start=1):
            if all(not str(cell).strip() for cell in row_data):
                issues.append(f"Row {row_idx} is completely empty")
        
        return issues
    
    def sync_csv_data_from_table(self):
        """Synchronize self.csv_data with current table contents in visual order"""
        if not self.csv_table.rowCount() or not self.csv_table.columnCount():
            return
        
        header = self.csv_table.horizontalHeader()
        col_count = self.csv_table.columnCount()
        
        # Get columns in visual order (respecting drag-and-drop reordering)
        visual_order = []
        for visual_pos in range(col_count):
            logical_pos = header.logicalIndex(visual_pos)
            visual_order.append(logical_pos)
        
        print(f"Saving with visual order: {visual_order}")
        
        # Collect headers in visual order
        headers = []
        for logical_col in visual_order:
            header_item = self.csv_table.horizontalHeaderItem(logical_col)
            headers.append(header_item.text() if header_item else f"Column {logical_col}")
        
        # Collect all rows in visual order
        data = [headers]
        for row in range(self.csv_table.rowCount()):
            row_data = []
            for logical_col in visual_order:
                item = self.csv_table.item(row, logical_col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        
        # Update csv_data
        self.csv_data = data
    
    def show_validation_dialog(self):
        """Show data validation results"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        # Sync data first
        self.sync_csv_data_from_table()
        
        # Validate
        issues = self.validate_csv_data()
        
        # Show results
        dialog = QDialog(self)
        dialog.setWindowTitle("Data Validation")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(400)
        
        layout = QVBoxLayout()
        
        if issues:
            # Show issues
            label = QLabel(f"<b>Found {len(issues)} issue(s):</b>")
            layout.addWidget(label)
            
            issues_text = QTextEdit()
            issues_text.setReadOnly(True)
            issues_text.setPlainText("\n".join(f"• {issue}" for issue in issues))
            layout.addWidget(issues_text)
        else:
            # No issues
            label = QLabel("<b>Data validation passed.</b><br><br>No issues found.")
            label.setStyleSheet("color: green; font-size: 14px;")
            layout.addWidget(label)
            
            info_text = QTextEdit()
            info_text.setReadOnly(True)
            info_text.setMaximumHeight(150)
            info_text.setPlainText(
                f"Total rows: {len(self.csv_data) - 1}\n"
                f"Total columns: {len(self.csv_data[0]) if self.csv_data else 0}\n"
                f"Headers: {', '.join(self.csv_data[0]) if self.csv_data else 'None'}"
            )
            layout.addWidget(info_text)
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)
        
        dialog.setLayout(layout)
        dialog.exec()
    
    def save_file(self):
        """Save current CSV file"""
        if not self.current_file:
            self.save_file_as()
            return
        
        try:
            # Sync data from table to csv_data
            self.sync_csv_data_from_table()
            
            # Validate data before saving
            issues = self.validate_csv_data()
            if issues:
                # Show validation warning
                reply = QMessageBox.question(
                    self, "Validation Issues",
                    f"Found {len(issues)} validation issue(s):\n\n" +
                    "\n".join(f"• {issue}" for issue in issues[:5]) +
                    (f"\n... and {len(issues) - 5} more" if len(issues) > 5 else "") +
                    "\n\nDo you want to save anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    return
            
            # Write to file with quotes around all non-empty fields
            with open(self.current_file, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC, lineterminator='\n')
                writer.writerows(self.csv_data)
            
            # Mark as saved (not modified)
            if self.current_file in self.modified_files:
                self.modified_files.remove(self.current_file)
            
            # Update cache
            self.file_data_cache[self.current_file] = self.csv_data.copy()
            
            self.update_file_tree_indicators()
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
            
            self.update_file_tree_indicators()
    
    def save_all_files(self):
        """Save all modified files"""
        if not self.modified_files:
            QMessageBox.information(self, "No Changes", "No files have been modified.")
            return
        
        # Save current file data to cache first
        if self.current_file and self.csv_data:
            self.sync_csv_data_from_table()
            self.file_data_cache[self.current_file] = self.csv_data.copy()
        
        saved_count = 0
        failed_files = []
        
        for file_path in list(self.modified_files):
            try:
                # Get data from cache
                if file_path in self.file_data_cache:
                    data_to_save = self.file_data_cache[file_path]
                elif file_path == self.current_file:
                    data_to_save = self.csv_data
                else:
                    continue
                
                # Write to file
                with open(file_path, 'w', encoding='utf-8', newline='') as f:
                    writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC, lineterminator='\n')
                    writer.writerows(data_to_save)
                
                saved_count += 1
                self.modified_files.discard(file_path)
                
            except Exception as e:
                failed_files.append((Path(file_path).name, str(e)))
        
        # Update indicators
        self.update_file_tree_indicators()
        
        # Show result
        if failed_files:
            error_msg = "\n".join([f"• {name}: {error}" for name, error in failed_files])
            QMessageBox.warning(
                self, "Save All - Partial Success",
                f"Saved {saved_count} file(s) successfully.\n\n"
                f"Failed to save {len(failed_files)} file(s):\n{error_msg}"
            )
        else:
            QMessageBox.information(
                self, "Save All Complete",
                f"Successfully saved {saved_count} file(s)."
            )
        
        self.status_bar.showMessage(f"Saved {saved_count} file(s)")
    
    def update_file_tree_indicators(self):
        """Update file tree to show modified indicators"""
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            file_path = item.data(0, Qt.ItemDataRole.UserRole)
            file_name = Path(file_path).name
            
            if file_path in self.modified_files:
                # Show asterisk for modified files
                item.setText(0, f"* {file_name}")
            else:
                # No asterisk for saved files
                item.setText(0, file_name)
    
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
            
            # Mark as modified
            if self.current_file:
                self.modified_files.add(self.current_file)
                self.file_data_cache[self.current_file] = self.csv_data.copy()
                self.update_file_tree_indicators()
            
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
            
            # Mark as modified
            if self.current_file:
                self.modified_files.add(self.current_file)
                self.file_data_cache[self.current_file] = self.csv_data.copy()
                self.update_file_tree_indicators()
            
            self.status_bar.showMessage(f"Inserted column: {column_name}")
    
    def delete_column(self):
        """Delete the selected column"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please select a column to delete.")
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
            
            # Mark as modified
            if self.current_file:
                self.modified_files.add(self.current_file)
                self.file_data_cache[self.current_file] = self.csv_data.copy()
                self.update_file_tree_indicators()
            
            self.status_bar.showMessage(f"Deleted column: {column_name}")
    
    def add_row(self):
        """Add a new empty row at the end"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        # Get number of columns
        col_count = self.csv_table.columnCount()
        if col_count == 0:
            QMessageBox.warning(self, "No Columns", "Cannot add row to a table with no columns.")
            return
        
        # Temporarily disconnect signal
        self.csv_table.itemChanged.disconnect(self.on_cell_changed)
        
        # Add row to table
        row_position = self.csv_table.rowCount()
        self.csv_table.insertRow(row_position)
        
        # Add empty cells
        for col in range(col_count):
            self.csv_table.setItem(row_position, col, QTableWidgetItem(""))
        
        # Add to csv_data
        self.csv_data.append([""] * col_count)
        
        # Reconnect signal
        self.csv_table.itemChanged.connect(self.on_cell_changed)
        
        # Mark as modified
        if self.current_file:
            self.modified_files.add(self.current_file)
            self.file_data_cache[self.current_file] = self.csv_data.copy()
            self.update_file_tree_indicators()
        
        self.status_bar.showMessage(f"Added row {row_position + 1}")
    
    def insert_row_before(self):
        """Insert a new empty row before the current row"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        current_row = self.csv_table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "No Selection", "Please select a row first.")
            return
        
        col_count = self.csv_table.columnCount()
        
        # Temporarily disconnect signal
        self.csv_table.itemChanged.disconnect(self.on_cell_changed)
        
        # Insert row in table
        self.csv_table.insertRow(current_row)
        
        # Add empty cells
        for col in range(col_count):
            self.csv_table.setItem(current_row, col, QTableWidgetItem(""))
        
        # Insert in csv_data (current_row + 1 because csv_data[0] is headers)
        self.csv_data.insert(current_row + 1, [""] * col_count)
        
        # Reconnect signal
        self.csv_table.itemChanged.connect(self.on_cell_changed)
        
        # Mark as modified
        if self.current_file:
            self.modified_files.add(self.current_file)
            self.file_data_cache[self.current_file] = self.csv_data.copy()
            self.update_file_tree_indicators()
        
        self.status_bar.showMessage(f"Inserted row at position {current_row + 1}")
    
    def delete_row(self):
        """Delete selected row(s)"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        # Get selected rows
        selected_rows = set()
        for item in self.csv_table.selectedItems():
            selected_rows.add(item.row())
        
        if not selected_rows:
            current_row = self.csv_table.currentRow()
            if current_row < 0:
                QMessageBox.warning(self, "No Selection", "Please select row(s) to delete.")
                return
            selected_rows.add(current_row)
        
        # Confirm deletion
        row_count = len(selected_rows)
        row_text = "row" if row_count == 1 else f"{row_count} rows"
        
        reply = QMessageBox.question(
            self, "Delete Row(s)",
            f"Are you sure you want to delete {row_text}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Temporarily disconnect signal
            self.csv_table.itemChanged.disconnect(self.on_cell_changed)
            
            # Sort rows in descending order to delete from bottom to top
            for row in sorted(selected_rows, reverse=True):
                # Remove from table
                self.csv_table.removeRow(row)
                
                # Remove from csv_data (row + 1 because csv_data[0] is headers)
                if row + 1 < len(self.csv_data):
                    del self.csv_data[row + 1]
            
            # Reconnect signal
            self.csv_table.itemChanged.connect(self.on_cell_changed)
            
            self.status_bar.showMessage(f"Deleted {row_count} {row_text}")
            
            # Mark as modified and update cache
            if self.current_file:
                self.modified_files.add(self.current_file)
                self.file_data_cache[self.current_file] = self.csv_data.copy()
                self.update_file_tree_indicators()
    
    def show_find_dialog(self):
        """Show find dialog"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Find")
        dialog.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        # Search text
        search_input = QLineEdit()
        search_input.setPlaceholderText("Enter text to find...")
        layout.addRow("Find:", search_input)
        
        # Options
        case_sensitive = QCheckBox("Case sensitive")
        layout.addRow(case_sensitive)
        
        whole_word = QCheckBox("Whole word")
        layout.addRow(whole_word)
        
        use_regex = QCheckBox("Regular expression")
        layout.addRow(use_regex)
        
        # Buttons
        button_layout = QHBoxLayout()
        find_all_btn = QPushButton("Find All")
        find_next_btn = QPushButton("Find Next")
        close_btn = QPushButton("Close")
        button_layout.addWidget(find_all_btn)
        button_layout.addWidget(find_next_btn)
        button_layout.addWidget(close_btn)
        layout.addRow(button_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        def do_find_all():
            self.perform_search(
                search_input.text(),
                case_sensitive.isChecked(),
                whole_word.isChecked(),
                use_regex.isChecked()
            )
        
        def do_find_next():
            if search_input.text():
                # Only perform search if no results exist yet
                if not self.search_results:
                    self.perform_search(
                        search_input.text(),
                        case_sensitive.isChecked(),
                        whole_word.isChecked(),
                        use_regex.isChecked()
                    )
                # Navigate to next result
                if self.search_results:
                    self.find_next()
        
        find_all_btn.clicked.connect(do_find_all)
        find_next_btn.clicked.connect(do_find_next)
        
        def close_and_clear():
            # Clear highlights when closing dialog
            if self.search_results:
                self.clear_search_highlights()
                self.search_results = []
                self.current_search_index = -1
            dialog.close()
        
        close_btn.clicked.connect(close_and_clear)
        
        # Allow Enter to find next
        search_input.returnPressed.connect(do_find_next)
        
        # Clear highlights when dialog is closed (X button)
        dialog.finished.connect(lambda: self.clear_search_highlights() if self.search_results else None)
        
        dialog.exec()
    
    def perform_search(self, search_text, case_sensitive, whole_word, use_regex):
        """Perform search and highlight results"""
        if not search_text:
            return
        
        try:
            # Clear previous highlights
            self.clear_search_highlights()
            self.search_results = []
        except Exception as e:
            print(f"Error in perform_search: {e}")
            QMessageBox.critical(self, "Search Error", f"An error occurred during search:\n{str(e)}")
            return
        
        # Prepare search pattern
        if use_regex:
            try:
                if case_sensitive:
                    pattern = re.compile(search_text)
                else:
                    pattern = re.compile(search_text, re.IGNORECASE)
            except re.error as e:
                QMessageBox.warning(self, "Invalid Regex", f"Invalid regular expression:\n{str(e)}")
                return
        else:
            # Escape special regex characters
            search_text_escaped = re.escape(search_text)
            if whole_word:
                search_text_escaped = r'\b' + search_text_escaped + r'\b'
            if case_sensitive:
                pattern = re.compile(search_text_escaped)
            else:
                pattern = re.compile(search_text_escaped, re.IGNORECASE)
        
        # Search through all cells
        for row in range(self.csv_table.rowCount()):
            for col in range(self.csv_table.columnCount()):
                item = self.csv_table.item(row, col)
                if item and item.text() and pattern.search(item.text()):
                    self.search_results.append((row, col))
                    # Highlight matching cell
                    try:
                        item.setBackground(QColor(255, 255, 0, 100))  # Light yellow
                    except Exception as e:
                        print(f"Error highlighting cell [{row}, {col}]: {e}")
        
        # Show results
        if self.search_results:
            self.current_search_index = 0
            self.highlight_current_result()
            self.status_bar.showMessage(f"Found {len(self.search_results)} matches")
        else:
            self.status_bar.showMessage("No matches found")
            QMessageBox.information(self, "No Results", f"No matches found for '{search_text}'")
    
    def clear_search_highlights(self):
        """Clear search highlights only from previously highlighted cells"""
        try:
            # Only clear highlights from cells that were in search results
            for row, col in self.search_results:
                item = self.csv_table.item(row, col)
                if item:
                    # Reset to default (transparent/no background)
                    item.setData(Qt.ItemDataRole.BackgroundRole, None)
        except Exception as e:
            print(f"Error clearing highlights: {e}")
    
    def highlight_current_result(self):
        """Highlight the current search result"""
        if not self.search_results or self.current_search_index < 0:
            return
        
        try:
            # Set flag to prevent clearing during navigation
            self.search_active = True
            
            # Reset all search results to light yellow
            for idx, (row, col) in enumerate(self.search_results):
                item = self.csv_table.item(row, col)
                if item:
                    if idx == self.current_search_index:
                        # Current result: Orange
                        item.setBackground(QColor(255, 165, 0, 150))
                    else:
                        # Other results: Light yellow
                        item.setBackground(QColor(255, 255, 0, 100))
            
            # Scroll to current result
            row, col = self.search_results[self.current_search_index]
            self.csv_table.setCurrentCell(row, col)
            item = self.csv_table.item(row, col)
            if item:
                self.csv_table.scrollToItem(item)
            
            self.status_bar.showMessage(
                f"Match {self.current_search_index + 1} of {len(self.search_results)}"
            )
            
            # Reset flag after a short delay
            QTimer.singleShot(100, lambda: setattr(self, 'search_active', False))
            
        except Exception as e:
            print(f"Error highlighting result: {e}")
            self.search_active = False
    
    def find_next(self):
        """Find next match"""
        if not self.search_results:
            QMessageBox.information(self, "No Results", "No search results. Use Find first.")
            return
        
        self.search_active = True
        self.current_search_index = (self.current_search_index + 1) % len(self.search_results)
        self.highlight_current_result()
    
    def find_previous(self):
        """Find previous match"""
        if not self.search_results:
            QMessageBox.information(self, "No Results", "No search results. Use Find first.")
            return
        
        self.search_active = True
        self.current_search_index = (self.current_search_index - 1) % len(self.search_results)
        self.highlight_current_result()
    
    def show_replace_dialog(self):
        """Show find and replace dialog"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please load a CSV file first.")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Find and Replace")
        dialog.setMinimumWidth(450)
        
        layout = QFormLayout()
        
        # Find text
        find_input = QLineEdit()
        find_input.setPlaceholderText("Enter text to find...")
        layout.addRow("Find:", find_input)
        
        # Replace text
        replace_input = QLineEdit()
        replace_input.setPlaceholderText("Enter replacement text...")
        layout.addRow("Replace with:", replace_input)
        
        # Options
        case_sensitive = QCheckBox("Case sensitive")
        layout.addRow(case_sensitive)
        
        whole_word = QCheckBox("Whole word")
        layout.addRow(whole_word)
        
        use_regex = QCheckBox("Regular expression")
        layout.addRow(use_regex)
        
        # Buttons
        button_layout = QHBoxLayout()
        replace_btn = QPushButton("Replace")
        replace_all_btn = QPushButton("Replace All")
        close_btn = QPushButton("Close")
        button_layout.addWidget(replace_btn)
        button_layout.addWidget(replace_all_btn)
        button_layout.addWidget(close_btn)
        layout.addRow(button_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        def do_replace():
            if not self.search_results or self.current_search_index < 0:
                # Perform search first
                self.perform_search(
                    find_input.text(),
                    case_sensitive.isChecked(),
                    whole_word.isChecked(),
                    use_regex.isChecked()
                )
                if not self.search_results:
                    return
            
            # Replace current match
            row, col = self.search_results[self.current_search_index]
            item = self.csv_table.item(row, col)
            if item:
                old_text = item.text()
                if use_regex:
                    try:
                        if case_sensitive.isChecked():
                            pattern = re.compile(find_input.text())
                        else:
                            pattern = re.compile(find_input.text(), re.IGNORECASE)
                        new_text = pattern.sub(replace_input.text(), old_text)
                    except re.error as e:
                        QMessageBox.warning(dialog, "Invalid Regex", f"Invalid regular expression:\n{str(e)}")
                        return
                else:
                    if case_sensitive.isChecked():
                        new_text = old_text.replace(find_input.text(), replace_input.text())
                    else:
                        # Case-insensitive replace
                        pattern = re.compile(re.escape(find_input.text()), re.IGNORECASE)
                        new_text = pattern.sub(replace_input.text(), old_text)
                
                item.setText(new_text)
            
            # Move to next match
            self.find_next()
        
        def do_replace_all():
            find_text = find_input.text()
            replace_text = replace_input.text()
            
            if not find_text:
                return
            
            # Perform search
            self.perform_search(
                find_text,
                case_sensitive.isChecked(),
                whole_word.isChecked(),
                use_regex.isChecked()
            )
            
            if not self.search_results:
                return
            
            # Confirm
            reply = QMessageBox.question(
                dialog, "Replace All",
                f"Replace {len(self.search_results)} occurrences?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                replaced_count = 0
                for row, col in self.search_results:
                    item = self.csv_table.item(row, col)
                    if item:
                        old_text = item.text()
                        if use_regex.isChecked():
                            try:
                                if case_sensitive.isChecked():
                                    pattern = re.compile(find_text)
                                else:
                                    pattern = re.compile(find_text, re.IGNORECASE)
                                new_text = pattern.sub(replace_text, old_text)
                            except re.error:
                                continue
                        else:
                            if case_sensitive.isChecked():
                                new_text = old_text.replace(find_text, replace_text)
                            else:
                                pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                                new_text = pattern.sub(replace_text, old_text)
                        
                        item.setText(new_text)
                        replaced_count += 1
                
                self.clear_search_highlights()
                self.search_results = []
                self.status_bar.showMessage(f"Replaced {replaced_count} occurrences")
                QMessageBox.information(dialog, "Replace Complete", f"Replaced {replaced_count} occurrences")
        
        replace_btn.clicked.connect(do_replace)
        replace_all_btn.clicked.connect(do_replace_all)
        
        def close_and_clear():
            # Clear highlights when closing dialog
            if self.search_results:
                self.clear_search_highlights()
                self.search_results = []
                self.current_search_index = -1
            dialog.close()
        
        close_btn.clicked.connect(close_and_clear)
        
        # Clear highlights when dialog is closed (X button)
        dialog.finished.connect(lambda: self.clear_search_highlights() if self.search_results else None)
        
        dialog.exec()
    

    
    def translate_with_google(self, text, source_lang, target_lang):
        if not DEEP_TRANSLATOR_AVAILABLE:
            raise Exception("deep-translator not available")
        translator = GoogleTranslator(source=source_lang.lower(), target=target_lang.lower())
        return translator.translate(text)
    

    def translate_with_mymemory(self, text, source_lang, target_lang):
        if not DEEP_TRANSLATOR_AVAILABLE:
            raise Exception("deep-translator not available")
        
        # MyMemory expects full locale codes (e.g., 'en-GB', 'ko-KR', 'ja-JP')
        # Map common 2-letter codes to full locale codes
        lang_map = {
            'en': 'en-GB',
            'ko': 'ko-KR',
            'ja': 'ja-JP',
            'zh': 'zh-CN',
            'es': 'es-ES',
            'fr': 'fr-FR',
            'de': 'de-DE',
            'it': 'it-IT',
            'pt': 'pt-PT',
            'ru': 'ru-RU',
            'ar': 'ar-SA',
            'hi': 'hi-IN',
            'th': 'th-TH',
            'vi': 'vi-VN',
            'id': 'id-ID',
            'tr': 'tr-TR',
            'pl': 'pl-PL',
            'nl': 'nl-NL',
            'sv': 'sv-SE',
            'da': 'da-DK',
            'fi': 'fi-FI',
            'no': 'nb-NO',
            'cs': 'cs-CZ',
            'hu': 'hu-HU',
            'ro': 'ro-RO',
            'uk': 'uk-UA',
            'el': 'el-GR',
            'he': 'he-IL',
        }
        
        # Convert to lowercase and get first 2 letters
        source_code = source_lang.lower()[:2]
        target_code = target_lang.lower()[:2]
        
        # Map to full locale codes
        source = lang_map.get(source_code, f"{source_code}-{source_code.upper()}")
        target = lang_map.get(target_code, f"{target_code}-{target_code.upper()}")
        
        translator = MyMemoryTranslator(source=source, target=target)
        return translator.translate(text)
    
    def translate_text(self, text, source_lang, target_lang):
        """
        Translate text using available services with centralized retry logic.
        Returns tuple of (translated_text, service_used).
        
        If tenacity is available, retries are handled by the decorator.
        Otherwise, falls back to simple retry logic.
        """
        if TENACITY_AVAILABLE:
            return self._translate_text_with_retry(text, source_lang, target_lang)
        else:
            return self._translate_text_simple_retry(text, source_lang, target_lang)
    
    def _translate_text_core(self, text, source_lang, target_lang):
        """Core translation logic that tries each service in priority order."""
        services = self.config['priority_order']
        last_exception = None
        
        for service in services:
            if service not in self.config['enabled_services']:
                continue
            try:
                if service == 'google':
                    result = self.translate_with_google(text, source_lang, target_lang)
                elif service == 'mymemory':
                    result = self.translate_with_mymemory(text, source_lang, target_lang)
                else:
                    continue
                self.log_translation(text, result, service, True, "")
                return (result, service)
            except Exception as e:
                self.log_translation(text, '', service, False, str(e))
                last_exception = e
                continue
        
        # All services failed
        raise last_exception if last_exception else Exception("All translation services failed")
    
    def _translate_text_with_retry(self, text, source_lang, target_lang):
        """Translation with tenacity retry decorator."""
        # Create a retry decorator dynamically based on config
        retry_decorator = retry(
            stop=stop_after_attempt(self.config['retry_count']),
            wait=wait_exponential(multiplier=1, min=1, max=10),
            reraise=True
        )
        
        # Apply decorator to core translation function
        retrying_translate = retry_decorator(self._translate_text_core)
        return retrying_translate(text, source_lang, target_lang)
    
    def _translate_text_simple_retry(self, text, source_lang, target_lang):
        """Fallback translation with simple retry logic when tenacity is not available."""
        max_retries = self.config['retry_count']
        last_exception = None
        
        for attempt in range(max_retries):
            try:
                return self._translate_text_core(text, source_lang, target_lang)
            except Exception as e:
                last_exception = e
                if attempt < max_retries - 1:
                    # Simple exponential backoff
                    wait_time = min(10, 2 ** attempt)
                    time.sleep(wait_time)
                continue
        
        # All retries exhausted
        raise last_exception if last_exception else Exception("All translation services failed")
    
    def log_translation(self, source_text, target_text, service, success, error_msg):
        self.translation_log.append({
            'timestamp': time.time(),
            'source': source_text,
            'target': target_text,
            'service': service,
            'success': success,
            'error': error_msg
        })
        if len(self.translation_log) > 1000:
            self.translation_log = self.translation_log[-1000:]
    

    
    def check_endpoint_health(self):
        self.endpoint_status = {}
        
        # Only mark Google/MyMemory as available if deep-translator is installed
        if DEEP_TRANSLATOR_AVAILABLE:
            self.endpoint_status['google'] = True
            self.endpoint_status['mymemory'] = True
        else:
            self.endpoint_status['google'] = False
            self.endpoint_status['mymemory'] = False
    
    def is_endpoint_disabled(self, endpoint):
        if endpoint not in self.circuit_breaker:
            return False
        failures, last_failure = self.circuit_breaker[endpoint]
        if failures >= self.config['circuit_breaker_threshold']:
            if time.time() - last_failure < self.config['circuit_breaker_timeout']:
                return True
            else:
                del self.circuit_breaker[endpoint]
        return False
    
    def record_endpoint_failure(self, endpoint):
        if endpoint not in self.circuit_breaker:
            self.circuit_breaker[endpoint] = [0, 0]
        self.circuit_breaker[endpoint][0] += 1
        self.circuit_breaker[endpoint][1] = time.time()
    
    def validate_translation_readiness(self):
        self.check_endpoint_health()
        available = any(self.endpoint_status.values())
        if not available:
            if not DEEP_TRANSLATOR_AVAILABLE:
                QMessageBox.warning(
                    self, 
                    "No Translation Services", 
                    "No translation services are available.\n\n"
                    "The 'deep-translator' package is not installed, so Google Translate "
                    "and MyMemory are unavailable.\n\n"
                    "Please install deep-translator: pip install deep-translator"
                )
            else:
                QMessageBox.warning(
                    self, 
                    "No Translation Services", 
                    "No translation services are available. Please check your internet connection or configure alternative services."
                )
            return False
        return True
    

    

    
    def translate_selected_cells(self, selected_items, source_lang, target_lang):
        """Translate selected cells"""
        if not self.validate_translation_readiness():
            return
        
        total_cells = len(selected_items)
        
        # Create progress dialog
        progress = QProgressDialog("Translating selected cells...", "Cancel", 0, total_cells, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        
        translated_count = 0
        failed_count = 0
        consecutive_failures = 0
        base_delay = self.config['base_delay']
        max_delay = 60.0
        current_delay = base_delay
        current_service = ""
        
        for idx, item in enumerate(selected_items):
            if progress.wasCanceled():
                break
            
            progress.setValue(idx)
            cell_text = item.text().strip()
            if not cell_text:
                continue
            
            # Translate (retry logic is now centralized in translate_text)
            try:
                result, service_used = self.translate_text(cell_text, source_lang, target_lang)
                current_service = service_used
                
                item.setText(result)
                translated_count += 1
                consecutive_failures = 0
                current_delay = max(base_delay, current_delay * 0.9)
                progress.setLabelText(f"Translating cell {idx + 1} of {total_cells}... (service: {current_service}, delay: {current_delay:.1f}s)")
                
                # Rate limiting with jitter
                time.sleep(current_delay + random.uniform(0, current_delay * 0.1))
            except Exception as e:
                failed_count += 1
                consecutive_failures += 1
                # Adaptive backoff for consecutive failures
                wait_time = min(max_delay, base_delay * (2 ** min(consecutive_failures - 1, 4)))
                current_delay = min(max_delay, current_delay * 1.5)
                progress.setLabelText(f"⚠️ Failed cell {idx + 1}. Waiting {wait_time:.1f}s...")
                time.sleep(wait_time)
        
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
                "Possible causes:\n"
                "• All translation services are unavailable\n"
                "• Network connection issues\n"
                "• Rate limiting on all services\n\n"
                "Check Settings > Translation Services Configuration."
            )
    
    def show_column_context_menu(self, position):
        """Show context menu for column header"""
        if not self.current_file:
            return
        
        column = self.csv_table.horizontalHeader().logicalIndexAt(position)
        if column < 0:
            return
        
        menu = QMenu()
        translate_action = QAction("Translate Entire Column with...", self)
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
        
        # Translation service
        service_combo = QComboBox()
        for service in TranslationService:
            service_combo.addItem(service.value.capitalize(), service.value)
        service_combo.setCurrentText(self.config['preferred_service'].capitalize())
        layout.addRow("Translation Service:", service_combo)
        
        # Use fallback
        use_fallback = QCheckBox("Use all available services as fallback")
        use_fallback.setChecked(True)
        layout.addRow(use_fallback)
        
        # Update source language when source column changes
        def update_source_lang():
            lang = detect_lang_code(source_combo.currentText())
            if lang:
                source_lang.setText(lang)
        source_combo.currentTextChanged.connect(update_source_lang)
        
        # Info label with status check
        info_label = QLabel("Checking translation services...")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow(info_label)
        
        def update_info():
            self.check_endpoint_health()
            available = [s for s in self.endpoint_status if self.endpoint_status[s]]
            if available:
                info_label.setText(f"Available services: {', '.join(available)}")
                info_label.setStyleSheet("color: green; font-size: 10px;")
            else:
                info_label.setText("No translation services available")
                info_label.setStyleSheet("color: red; font-size: 10px;")
        
        update_info()
        
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
            service = service_combo.currentData()
            fallback = use_fallback.isChecked()
            
            if not src_lang or not tgt_lang:
                QMessageBox.warning(self, "Invalid Input", "Please enter both source and target languages.")
                return
            
            # Update preferred service but don't modify priority_order permanently
            self.config['preferred_service'] = service
            self.save_config()
            
            # Temporarily adjust priority for this translation only
            if not fallback:
                # Save original priority order
                original_priority = self.config['priority_order'].copy()
                original_enabled = self.config['enabled_services'].copy()
                
                # Temporarily set to single service
                self.config['priority_order'] = [service]
                self.config['enabled_services'] = [service]
                
                # Start translation
                self.translate_column(source_col, target_column, src_lang, tgt_lang)
                
                # Restore original settings
                self.config['priority_order'] = original_priority
                self.config['enabled_services'] = original_enabled
            else:
                # Use fallback - put selected service first
                temp_priority = [service] + [s for s in self.config['priority_order'] if s != service]
                original_priority = self.config['priority_order'].copy()
                self.config['priority_order'] = temp_priority
                
                # Start translation
                self.translate_column(source_col, target_column, src_lang, tgt_lang)
                
                # Restore original priority
                self.config['priority_order'] = original_priority
    
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
        
        # Translation service
        service_combo = QComboBox()
        for service in TranslationService:
            service_combo.addItem(service.value.capitalize(), service.value)
        service_combo.setCurrentText(self.config['preferred_service'].capitalize())
        layout.addRow("Translation Service:", service_combo)
        
        # Use fallback
        use_fallback = QCheckBox("Use all available services as fallback")
        use_fallback.setChecked(True)
        layout.addRow(use_fallback)
        
        # Info label
        status_label = QLabel("Checking translation services...")
        status_label.setWordWrap(True)
        status_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow(status_label)
        
        def update_status():
            self.check_endpoint_health()
            available = [s for s in self.endpoint_status if self.endpoint_status[s]]
            if available:
                status_label.setText(f"Available services: {', '.join(available)}")
                status_label.setStyleSheet("color: green; font-size: 10px;")
            else:
                status_label.setText("No translation services available")
                status_label.setStyleSheet("color: red; font-size: 10px;")
        
        update_status()
        
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
            service = service_combo.currentData()
            fallback = use_fallback.isChecked()
            
            if not src_lang or not tgt_lang:
                QMessageBox.warning(self, "Invalid Input", "Please enter both source and target languages.")
                return
            
            # Update config
            self.config['preferred_service'] = service
            if fallback:
                self.config['priority_order'] = [service] + [s for s in self.config['priority_order'] if s != service]
            else:
                self.config['priority_order'] = [service]
            self.save_config()
            
            # Start translation
            self.translate_selected_cells(selected_items, src_lang, tgt_lang)
    
    def translate_column(self, source_col, target_col, source_lang, target_lang):
        """Translate all cells from source column to target column"""
        if not self.validate_translation_readiness():
            return
        
        row_count = self.csv_table.rowCount()
        
        # Create progress dialog
        progress = QProgressDialog("Translating...", "Cancel", 0, row_count, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        
        translated_count = 0
        failed_count = 0
        consecutive_failures = 0
        base_delay = self.config['base_delay']
        max_delay = 60.0
        current_delay = base_delay
        current_service = ""
        paused = False
        
        pause_btn = QPushButton("Pause")
        progress.setCancelButton(pause_btn)
        pause_btn.clicked.connect(lambda: setattr(progress, 'paused', not getattr(progress, 'paused', False)))
        
        for row in range(row_count):
            if progress.wasCanceled():
                break
            
            while getattr(progress, 'paused', False):
                QApplication.processEvents()
                time.sleep(0.1)
            
            progress.setValue(row)
            
            # Get source text
            source_item = self.csv_table.item(row, source_col)
            if not source_item or not source_item.text().strip():
                continue
            
            source_text = source_item.text()
            
            # Translate (retry logic is now centralized in translate_text)
            try:
                result, service_used = self.translate_text(source_text, source_lang, target_lang)
                current_service = service_used
                
                # Update target cell
                target_item = self.csv_table.item(row, target_col)
                if target_item:
                    target_item.setText(result)
                else:
                    self.csv_table.setItem(row, target_col, QTableWidgetItem(result))
                
                translated_count += 1
                consecutive_failures = 0
                current_delay = max(base_delay, current_delay * 0.9)
                progress.setLabelText(f"Translating row {row + 1} of {row_count}... (service: {current_service}, delay: {current_delay:.1f}s)")
                
                # Rate limiting with jitter
                time.sleep(current_delay + random.uniform(0, current_delay * 0.1))
            except Exception as e:
                failed_count += 1
                consecutive_failures += 1
                # Adaptive backoff for consecutive failures
                wait_time = min(max_delay, base_delay * (2 ** min(consecutive_failures - 1, 4)))
                current_delay = min(max_delay, current_delay * 1.5)
                progress.setLabelText(f"⚠️ Failed row {row + 1}. Waiting {wait_time:.1f}s...")
                time.sleep(wait_time)
        
        progress.setValue(row_count)
        
        # Show summary
        if translated_count > 0:
            message = f"Translation complete!\n\nTranslated: {translated_count} rows"
            if failed_count > 0:
                message += f"\nFailed: {failed_count} rows"
            QMessageBox.information(self, "Translation Complete", message)
            self.status_bar.showMessage(f"Translated {translated_count} rows")
        else:
            QMessageBox.warning(
                self, "Translation Failed", 
                "No rows were translated.\n\n"
                "Possible causes:\n"
                "• All translation services are unavailable\n"
                "• Network connection issues\n"
                "• Rate limiting on all services\n\n"
                "Check Settings > Translation Services Configuration."
            )
    
    def show_translation_config_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Translation Services Configuration")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(400)
        
        layout = QVBoxLayout()
        
        # Enabled services
        enabled_group = QGroupBox("Enabled Services")
        enabled_layout = QVBoxLayout()
        
        service_checks = {}
        for service in TranslationService:
            check = QCheckBox(service.value.capitalize())
            check.setChecked(service.value in self.config['enabled_services'])
            service_checks[service.value] = check
            enabled_layout.addWidget(check)
        
        enabled_group.setLayout(enabled_layout)
        layout.addWidget(enabled_group)
        
        # Priority order
        priority_group = QGroupBox("Priority Order (drag to reorder)")
        priority_layout = QVBoxLayout()
        
        priority_list = QListWidget()
        # Get valid services from the enum
        valid_services = [s.value for s in TranslationService]
        # Filter priority order to only include valid services
        for service in self.config['priority_order']:
            if service in valid_services:
                item = QListWidgetItem(service.capitalize())
                item.setData(Qt.ItemDataRole.UserRole, service)
                priority_list.addItem(item)
        
        priority_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        priority_layout.addWidget(priority_list)
        
        priority_group.setLayout(priority_layout)
        layout.addWidget(priority_group)
        
        # Settings
        settings_group = QGroupBox("Settings")
        settings_layout = QFormLayout()
        
        timeout_spin = QSpinBox()
        timeout_spin.setRange(5, 120)
        timeout_spin.setValue(self.config['request_timeout'])
        settings_layout.addRow("Request Timeout (s):", timeout_spin)
        
        retry_spin = QSpinBox()
        retry_spin.setRange(1, 10)
        retry_spin.setValue(self.config['retry_count'])
        settings_layout.addRow("Retry Count:", retry_spin)
        
        delay_spin = QDoubleSpinBox()
        delay_spin.setRange(0.1, 30.0)
        delay_spin.setValue(self.config['base_delay'])
        settings_layout.addRow("Base Delay (s):", delay_spin)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        # Set the dialog layout
        dialog.setLayout(layout)
        
        save_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.config['enabled_services'] = [s for s in service_checks if service_checks[s].isChecked()]
            self.config['priority_order'] = [priority_list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(priority_list.count())]
            self.config['request_timeout'] = timeout_spin.value()
            self.config['retry_count'] = retry_spin.value()
            self.config['base_delay'] = delay_spin.value()
            self.save_config()
    
    def show_translation_log(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Translation Log")
        dialog.setMinimumWidth(800)
        dialog.setMinimumHeight(600)
        
        layout = QVBoxLayout()
        
        # Summary statistics section
        stats_group = QGroupBox("Summary Statistics")
        stats_layout = QVBoxLayout()
        
        stats = self.get_translation_stats()
        overall_stats = QLabel(
            f"<b>Overall:</b> {stats['total']} translations | "
            f"Success: {stats['successful']} ({stats['success_rate']:.1f}%) | "
            f"Failed: {stats['failed']}"
        )
        stats_layout.addWidget(overall_stats)
        
        # Per-service statistics
        if stats['by_service']:
            service_stats_text = "<b>By Service:</b><br>"
            for service, service_data in stats['by_service'].items():
                service_stats_text += (
                    f"&nbsp;&nbsp;• {service}: {service_data['total']} total, "
                    f"{service_data['successful']} success, "
                    f"{service_data['failed']} failed "
                    f"({service_data['success_rate']:.1f}%)<br>"
                )
            service_stats_label = QLabel(service_stats_text)
            stats_layout.addWidget(service_stats_label)
        
        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)
        
        # Log entries list
        log_label = QLabel("Recent Translations (last 100):")
        layout.addWidget(log_label)
        
        log_list = QListWidget()
        log_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        
        # Populate log list with recent entries
        for entry in reversed(self.translation_log[-100:]):
            time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(entry['timestamp']))
            status = "✓" if entry['success'] else "✗"
            source_preview = entry['source'][:40] + "..." if len(entry['source']) > 40 else entry['source']
            target_preview = entry['target'][:40] + "..." if len(entry['target']) > 40 else entry['target']
            
            item_text = f"{status} {time_str} [{entry['service']}] {source_preview} → {target_preview}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, entry)  # Store full entry data
            
            # Color code by status
            if entry['success']:
                item.setForeground(QColor(0, 128, 0))  # Green for success
            else:
                item.setForeground(QColor(200, 0, 0))  # Red for failure
            
            log_list.addItem(item)
        
        layout.addWidget(log_list)
        
        # Details section
        details_group = QGroupBox("Selected Entry Details")
        details_layout = QVBoxLayout()
        details_text = QTextEdit()
        details_text.setReadOnly(True)
        details_text.setMaximumHeight(150)
        details_text.setPlainText("Select an entry to view details...")
        details_layout.addWidget(details_text)
        details_group.setLayout(details_layout)
        layout.addWidget(details_group)
        
        # Update details when selection changes
        def update_details():
            selected_items = log_list.selectedItems()
            if selected_items:
                entry = selected_items[0].data(Qt.ItemDataRole.UserRole)
                time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(entry['timestamp']))
                details = f"Timestamp: {time_str}\n"
                details += f"Service: {entry['service']}\n"
                details += f"Status: {'SUCCESS' if entry['success'] else 'FAILED'}\n"
                details += f"Source Text: {entry['source']}\n"
                details += f"Target Text: {entry['target']}\n"
                if not entry['success'] and entry['error']:
                    details += f"\nError Details:\n{entry['error']}"
                details_text.setPlainText(details)
            else:
                details_text.setPlainText("Select an entry to view details...")
        
        log_list.itemSelectionChanged.connect(update_details)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        copy_details_btn = QPushButton("Copy Error Details")
        copy_details_btn.setToolTip("Copy error details of selected failed entry to clipboard")
        def copy_error_details():
            selected_items = log_list.selectedItems()
            if selected_items:
                entry = selected_items[0].data(Qt.ItemDataRole.UserRole)
                if not entry['success']:
                    time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(entry['timestamp']))
                    error_info = f"Translation Error Report\n"
                    error_info += f"{'=' * 50}\n"
                    error_info += f"Timestamp: {time_str}\n"
                    error_info += f"Service: {entry['service']}\n"
                    error_info += f"Source Language: (from translation context)\n"
                    error_info += f"Target Language: (from translation context)\n"
                    error_info += f"Source Text: {entry['source']}\n"
                    error_info += f"Error Message: {entry['error']}\n"
                    error_info += f"{'=' * 50}\n"
                    
                    clipboard = QApplication.clipboard()
                    clipboard.setText(error_info)
                    QMessageBox.information(dialog, "Copied", "Error details copied to clipboard.")
                else:
                    QMessageBox.information(dialog, "No Error", "Selected entry was successful (no error to copy).")
            else:
                QMessageBox.warning(dialog, "No Selection", "Please select a log entry first.")
        
        copy_details_btn.clicked.connect(copy_error_details)
        button_layout.addWidget(copy_details_btn)
        
        export_json_btn = QPushButton("Export to JSON")
        export_json_btn.setToolTip("Export full log to JSON file")
        export_json_btn.clicked.connect(lambda: self.export_translation_log('json'))
        button_layout.addWidget(export_json_btn)
        
        export_csv_btn = QPushButton("Export to CSV")
        export_csv_btn.setToolTip("Export log to CSV file")
        export_csv_btn.clicked.connect(lambda: self.export_translation_log('csv'))
        button_layout.addWidget(export_csv_btn)
        
        clear_btn = QPushButton("Clear Log")
        clear_btn.clicked.connect(lambda: self.clear_translation_log() or dialog.close())
        button_layout.addWidget(clear_btn)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        dialog.setLayout(layout)
        dialog.exec()
    
    def get_translation_stats(self):
        """Calculate detailed translation statistics including per-service breakdown"""
        total = len(self.translation_log)
        successful = sum(1 for e in self.translation_log if e['success'])
        failed = total - successful
        success_rate = (successful / total * 100) if total > 0 else 0
        
        # Per-service statistics
        by_service = {}
        for entry in self.translation_log:
            service = entry['service']
            if service not in by_service:
                by_service[service] = {'total': 0, 'successful': 0, 'failed': 0}
            
            by_service[service]['total'] += 1
            if entry['success']:
                by_service[service]['successful'] += 1
            else:
                by_service[service]['failed'] += 1
        
        # Calculate success rate per service
        for service in by_service:
            service_total = by_service[service]['total']
            service_successful = by_service[service]['successful']
            by_service[service]['success_rate'] = (service_successful / service_total * 100) if service_total > 0 else 0
        
        return {
            'total': total,
            'successful': successful,
            'failed': failed,
            'success_rate': success_rate,
            'by_service': by_service
        }
    
    def export_translation_log(self, format_type='json'):
        """Export translation log to JSON or CSV format"""
        if format_type == 'json':
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Export Translation Log", "", "JSON Files (*.json);;All Files (*)"
            )
            if file_path:
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(self.translation_log, f, indent=2, ensure_ascii=False)
                    QMessageBox.information(self, "Export Success", f"Log exported to {file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "Export Failed", f"Failed to export log:\n{str(e)}")
        
        elif format_type == 'csv':
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Export Translation Log", "", "CSV Files (*.csv);;All Files (*)"
            )
            if file_path:
                try:
                    with open(file_path, 'w', encoding='utf-8', newline='') as f:
                        writer = csv.writer(f)
                        # Write header
                        writer.writerow(['Timestamp', 'Service', 'Status', 'Source Text', 'Target Text', 'Error'])
                        
                        # Write data
                        for entry in self.translation_log:
                            time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(entry['timestamp']))
                            status = 'SUCCESS' if entry['success'] else 'FAILED'
                            writer.writerow([
                                time_str,
                                entry['service'],
                                status,
                                entry['source'],
                                entry['target'],
                                entry.get('error', '')
                            ])
                    QMessageBox.information(self, "Export Success", f"Log exported to {file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "Export Failed", f"Failed to export log:\n{str(e)}")
    
    def clear_translation_log(self):
        self.translation_log = []
        QMessageBox.information(self, "Log Cleared", "Translation log has been cleared.")



def main():
    app = QApplication(sys.argv)
    window = CSVEditorWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()