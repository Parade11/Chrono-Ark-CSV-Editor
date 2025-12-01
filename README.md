# CSV Editor - For Chrono Ark

A modern Python desktop application for editing CSV files, designed to simplify translation workflows.

## Features

- **Import CSV Files**: Manually import CSV files into the application
- **File Tree View**: View all imported CSV files in the left sidebar
- **Editable Table**: View and edit CSV content in a spreadsheet-like interface
- **Easy Navigation**: Click any CSV file in the tree to load it instantly
- **Manual Editing**: Double-click any cell to edit its content
- **Save Functionality**: Save changes with Ctrl+S or through the File menu
- **Remove Files**: Remove files from the list when no longer needed
- **Modern UI**: Clean, professional interface built with PyQt6
- **Multiple Translation Services**: Support for Google Translate and MyMemory with automatic fallback
- **Configurable Translation Settings**: Customize timeouts, retries, and service priorities

## Installation

1. Install Python 3.8 or higher
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python csv_editor.py
```

2. Import CSV files:
   - Click `File > Import CSV...` or press `Ctrl+O`
   - Select one or multiple CSV files to import
   
3. The left sidebar shows all imported CSV files

4. Click on any CSV file to view its contents in the table

5. Double-click any cell to edit it

6. Save your changes:
   - Press `Ctrl+S` to save
   - Or use `File > Save` from the menu
   - Use `File > Save As` to save with a new name

7. Remove files from the list:
   - Select a file in the tree and press `Delete`
   - Or use `File > Remove from List`

## Keyboard Shortcuts

- `Ctrl+O` - Import CSV file(s)
- `Ctrl+S` - Save current file
- `Ctrl+Shift+S` - Save As
- `Ctrl+Shift+C` - Add new column
- `Ctrl+Shift+D` - Delete selected column
- `Delete` - Remove selected file from list
- `Ctrl+Q` - Exit application

## Translation Workflow

This tool is designed for translation work:
1. Import your language CSV files (like LangDataDB.csv, LangSystemDB.csv)
2. Click on a file to view its contents
3. Edit translations directly in the table
4. **Add new language columns** if needed (e.g., add "English" column to a Chinese-only CSV)
5. Save changes and continue working
6. All changes are preserved in the original CSV format
7. Import additional files as needed

### Adding New Language Columns

If your CSV only has Chinese and you want to add English translations:
1. Load the CSV file
2. Go to `Edit > Add Column...` (or press `Ctrl+Shift+C`)
3. Enter the column name (e.g., "English")
4. The new column will be added at the end
5. Start filling in translations
6. Save the file

You can also:
- **Insert a column** at a specific position: Select a column, then `Edit > Insert Column Before...`
- **Delete a column**: Select a column, then `Edit > Delete Column` (or `Ctrl+Shift+D`)

### Translating Columns with Multiple Services

Automatically translate content from one column to another using multiple translation services:

1. **Translate a column**:
   - Right-click on the target column header (e.g., "Korean")
   - Select "Translate Column..."
   - Choose source column (e.g., "English")
   - Enter source language code (e.g., "EN")
   - Enter target language code (e.g., "KO")
   - Select preferred translation service or use automatic fallback
   - Click "Translate"

2. **Language codes**:
   - EN = English
   - ZH = Chinese
   - JA = Japanese
   - KO = Korean
   - DE = German
   - FR = French
   - ES = Spanish
   - And more...

#### Configuring Translation Services

To configure translation services:
1. Go to `Settings > Translation Services Configuration`
2. Enable/disable specific services
3. Set priority order for fallback
4. Adjust timeouts and retry settings

The automatic fallback mechanism tries services in priority order.

## Requirements

- Python 3.8+
- PyQt6 6.4.0+
- requests 2.28.0+ (for translation feature)
- chardet 5.0.0+
- deep-translator 1.11.4+ (for alternative translation services)
- tenacity 8.2.0+ (for retry logic)
- urllib3 2.0.0+ (for HTTP handling)
