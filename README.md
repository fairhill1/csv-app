# GridView v0.1

A lightweight spreadsheet application built with Rust, egui, and CSV support.

## Features

- **Open and Edit CSV Files**: Load CSV files and edit them in a spreadsheet-like interface
- **Column Resizing**: Drag column borders to resize columns
- **Cell Editing**: Click any cell to edit its content
- **Text Clipping**: Long text is clipped to cell boundaries
- **Add Rows/Columns**: Dynamically expand your spreadsheet
- **Save Changes**: Save back to CSV format
- **New Spreadsheets**: Create new spreadsheets from scratch
- **Context Menu**: Right-click cells to clear content
- **Spreadsheet-style Headers**: Columns labeled A, B, C... and rows numbered 1, 2, 3...

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

## Creating macOS App Bundle

The project includes a pre-built `GridView.app` bundle. To use it:

1. **Copy to Applications**:
   ```bash
   cp -r GridView.app /Applications/
   ```

2. **Add Custom Icon** (optional):
   - Create or download a 1024x1024 PNG icon
   - Run the icon creation script:
     ```bash
     ./create-icon.sh your-icon.png
     ```
   - The icon will be automatically added to `GridView.app`

3. **Rebuild the app bundle** (if you modify the code):
   ```bash
   cargo build --release
   cp target/release/csv-app GridView.app/Contents/MacOS/GridView
   ```

The app will launch without a terminal window and can be opened by double-clicking.

## License

MIT
