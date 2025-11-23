# CSV Spreadsheet App v0.1

A lightweight spreadsheet application built with Rust, egui, and CSV support.

## Features

- **Open and Edit CSV Files**: Load CSV files and edit them in a spreadsheet-like interface
- **Cell Editing**: Click any cell to edit its content
- **Add Rows/Columns**: Dynamically expand your spreadsheet
- **Save Changes**: Save back to CSV format
- **New Spreadsheets**: Create new spreadsheets from scratch
- **Context Menu**: Right-click cells to clear content

## Building

```bash
cargo build --release
```

## Running

```bash
cargo run --release
```

## Usage

### File Menu
- **New**: Create a new blank spreadsheet (10x20 grid)
- **Open CSV**: Load an existing CSV file
- **Save**: Save to the current file
- **Save As...**: Save to a new file

### Edit Menu
- **Add Row**: Add a new row at the bottom
- **Add Column**: Add a new column on the right

### Editing Cells
- Click on any cell to start editing
- Press Enter or click outside to confirm changes
- Right-click a cell and select "Clear" to empty it

## Sample File

A sample CSV file (`sample.csv`) is included for testing.

## Dependencies

- eframe 0.33
- egui 0.33
- csv 1.4
- rfd 0.15

## License

MIT
