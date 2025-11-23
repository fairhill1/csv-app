use eframe::egui;
use egui_extras::{TableBuilder, Column};
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Debug, Clone, PartialEq)]
enum Selection {
    None,
    CellRange { start: (usize, usize), end: (usize, usize) },
    Column(usize),
    Row(usize),
}

fn main() -> eframe::Result<()> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([1200.0, 800.0]),
        ..Default::default()
    };

    eframe::run_native(
        "CSV Spreadsheet",
        options,
        Box::new(|_cc| Ok(Box::new(SpreadsheetApp::default()))),
    )
}

struct SpreadsheetApp {
    data: Vec<Vec<String>>,
    file_path: Option<PathBuf>,
    editing_cell: Option<(usize, usize)>,
    edit_buffer: String,
    column_widths: HashMap<usize, f32>,
    default_column_width: f32,
    selection: Selection,
    drag_start: Option<(usize, usize)>,
    clipboard: arboard::Clipboard,
    undo_stack: Vec<Vec<Vec<String>>>,
    redo_stack: Vec<Vec<Vec<String>>>,
}

impl Default for SpreadsheetApp {
    fn default() -> Self {
        Self {
            data: vec![vec![String::new(); 10]; 20],
            file_path: None,
            editing_cell: None,
            edit_buffer: String::new(),
            column_widths: HashMap::new(),
            default_column_width: 120.0,
            selection: Selection::None,
            drag_start: None,
            clipboard: arboard::Clipboard::new().unwrap(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
        }
    }
}

impl SpreadsheetApp {
    fn col_index_to_letter(idx: usize) -> String {
        let mut result = String::new();
        let mut num = idx + 1;

        while num > 0 {
            num -= 1;
            result.insert(0, (b'A' + (num % 26) as u8) as char);
            num /= 26;
        }

        result
    }

    fn normalize_data(&mut self) {
        let max_cols = self.data.iter().map(|r| r.len()).max().unwrap_or(0);
        for row in &mut self.data {
            if row.len() < max_cols {
                row.resize(max_cols, String::new());
            }
        }
    }

    fn get_column_width(&self, col_idx: usize) -> f32 {
        *self.column_widths.get(&col_idx).unwrap_or(&self.default_column_width)
    }

    fn clear_selection(&mut self) {
        match &self.selection {
            Selection::None => {}
            Selection::CellRange { start, end } => {
                let (r1, c1) = *start;
                let (r2, c2) = *end;
                let (min_r, max_r) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
                let (min_c, max_c) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
                for row_idx in min_r..=max_r {
                    if row_idx < self.data.len() {
                        for col_idx in min_c..=max_c {
                            if col_idx < self.data[row_idx].len() {
                                self.data[row_idx][col_idx].clear();
                            }
                        }
                    }
                }
            }
            Selection::Column(col_idx) => {
                for row in &mut self.data {
                    if *col_idx < row.len() {
                        row[*col_idx].clear();
                    }
                }
            }
            Selection::Row(row_idx) => {
                if *row_idx < self.data.len() {
                    for cell in &mut self.data[*row_idx] {
                        cell.clear();
                    }
                }
            }
        }
    }

    fn load_csv(&mut self, path: PathBuf) {
        match csv::Reader::from_path(&path) {
            Ok(mut reader) => {
                let mut data = Vec::new();

                // Add headers as first row
                if let Ok(headers) = reader.headers() {
                    let header_row: Vec<String> = headers.iter().map(|s| s.to_string()).collect();
                    data.push(header_row);
                }

                // Add data rows
                for result in reader.records() {
                    if let Ok(record) = result {
                        let row: Vec<String> = record.iter().map(|s| s.to_string()).collect();
                        data.push(row);
                    }
                }

                self.data = data;
                // Normalize immediately to ensure rectangular structure
                self.normalize_data();
                self.file_path = Some(path);
            }
            Err(e) => {
                eprintln!("Error loading CSV: {}", e);
            }
        }
    }

    fn save_csv(&self, path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
        let mut writer = csv::Writer::from_path(path)?;

        for row in &self.data {
            writer.write_record(row)?;
        }

        writer.flush()?;
        Ok(())
    }

    fn add_row(&mut self) {
        let cols = self.data.first().map(|r| r.len()).unwrap_or(10);
        self.data.push(vec![String::new(); cols]);
    }

    fn add_column(&mut self) {
        if self.data.is_empty() {
            self.data.push(vec![String::new()]);
        } else {
            for row in &mut self.data {
                row.push(String::new());
            }
        }
    }

    fn insert_row_at(&mut self, row_idx: usize) {
        let cols = self.data.first().map(|r| r.len()).unwrap_or(10);
        self.data.insert(row_idx, vec![String::new(); cols]);

        // Adjust editing cell index if after inserted row
        if let Some((editing_row, editing_col)) = self.editing_cell {
            if editing_row >= row_idx {
                self.editing_cell = Some((editing_row + 1, editing_col));
            }
        }
    }

    fn insert_column_at(&mut self, col_idx: usize) {
        if self.data.is_empty() {
            self.data.push(vec![String::new()]);
        } else {
            for row in &mut self.data {
                row.insert(col_idx, String::new());
            }
        }

        // Adjust editing cell index if after inserted column
        if let Some((editing_row, editing_col)) = self.editing_cell {
            if editing_col >= col_idx {
                self.editing_cell = Some((editing_row, editing_col + 1));
            }
        }

        // Shift column widths for columns at or after the inserted one
        let mut new_widths = HashMap::new();
        for (&idx, &width) in &self.column_widths {
            if idx >= col_idx {
                new_widths.insert(idx + 1, width);
            } else {
                new_widths.insert(idx, width);
            }
        }
        self.column_widths = new_widths;
    }

    fn delete_row(&mut self, row_idx: usize) {
        if row_idx < self.data.len() {
            self.data.remove(row_idx);
            // Clear editing state if we're editing the deleted row
            if let Some((editing_row, _)) = self.editing_cell {
                if editing_row == row_idx {
                    self.editing_cell = None;
                } else if editing_row > row_idx {
                    // Adjust editing cell index if after deleted row
                    self.editing_cell = Some((editing_row - 1, self.editing_cell.unwrap().1));
                }
            }
        }
    }

    fn delete_column(&mut self, col_idx: usize) {
        for row in &mut self.data {
            if col_idx < row.len() {
                row.remove(col_idx);
            }
        }
        // Clear editing state if we're editing the deleted column
        if let Some((editing_row, editing_col)) = self.editing_cell {
            if editing_col == col_idx {
                self.editing_cell = None;
            } else if editing_col > col_idx {
                // Adjust editing cell index if after deleted column
                self.editing_cell = Some((editing_row, editing_col - 1));
            }
        }
        // Remove column width setting
        self.column_widths.remove(&col_idx);
        // Shift column widths for columns after the deleted one
        let mut new_widths = HashMap::new();
        for (&idx, &width) in &self.column_widths {
            if idx > col_idx {
                new_widths.insert(idx - 1, width);
            } else {
                new_widths.insert(idx, width);
            }
        }
        self.column_widths = new_widths;
    }

    fn save_undo_state(&mut self) {
        self.undo_stack.push(self.data.clone());
        self.redo_stack.clear();
        // Limit undo stack to 50 entries
        if self.undo_stack.len() > 50 {
            self.undo_stack.remove(0);
        }
    }

    fn undo(&mut self) {
        if let Some(prev_state) = self.undo_stack.pop() {
            self.redo_stack.push(self.data.clone());
            self.data = prev_state;
        }
    }

    fn redo(&mut self) {
        if let Some(next_state) = self.redo_stack.pop() {
            self.undo_stack.push(self.data.clone());
            self.data = next_state;
        }
    }

    fn copy_selection(&mut self) {
        let text = self.get_selection_as_text();
        if !text.is_empty() {
            let _ = self.clipboard.set_text(text);
        }
    }

    fn cut_selection(&mut self) {
        self.save_undo_state();
        let text = self.get_selection_as_text();
        if !text.is_empty() {
            let _ = self.clipboard.set_text(text);
            self.clear_selection();
        }
    }

    fn get_selection_as_text(&self) -> String {
        match &self.selection {
            Selection::None => String::new(),
            Selection::CellRange { start, end } => {
                let (r1, c1) = *start;
                let (r2, c2) = *end;
                let (min_r, max_r) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
                let (min_c, max_c) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };

                let mut rows = Vec::new();
                for row_idx in min_r..=max_r {
                    if row_idx < self.data.len() {
                        let mut cells = Vec::new();
                        for col_idx in min_c..=max_c {
                            if col_idx < self.data[row_idx].len() {
                                cells.push(self.data[row_idx][col_idx].clone());
                            } else {
                                cells.push(String::new());
                            }
                        }
                        rows.push(cells.join("\t"));
                    }
                }
                rows.join("\n")
            }
            Selection::Column(col_idx) => {
                let mut cells = Vec::new();
                for row in &self.data {
                    if *col_idx < row.len() {
                        cells.push(row[*col_idx].clone());
                    } else {
                        cells.push(String::new());
                    }
                }
                cells.join("\n")
            }
            Selection::Row(row_idx) => {
                if *row_idx < self.data.len() {
                    self.data[*row_idx].join("\t")
                } else {
                    String::new()
                }
            }
        }
    }

    fn paste_text(&mut self, text: &str) {
        // Determine starting position based on selection
        let (start_row, start_col) = match &self.selection {
            Selection::CellRange { start, .. } => *start,
            Selection::Row(r) => (*r, 0),
            Selection::Column(c) => (0, *c),
            Selection::None => (0, 0),
        };

        let lines: Vec<&str> = text.lines().collect();

        // Calculate max columns needed
        let max_cols_needed = self.data.iter().map(|r| r.len()).max().unwrap_or(10);

        for (row_offset, line) in lines.iter().enumerate() {
            let row_idx = start_row + row_offset;
            let cells: Vec<&str> = line.split('\t').collect();

            // Ensure we have enough rows
            while row_idx >= self.data.len() {
                self.data.push(vec![String::new(); max_cols_needed]);
            }

            for (col_offset, cell_text) in cells.iter().enumerate() {
                let col_idx = start_col + col_offset;

                // Ensure we have enough columns
                while col_idx >= self.data[row_idx].len() {
                    self.data[row_idx].push(String::new());
                }

                self.data[row_idx][col_idx] = cell_text.to_string();
            }
        }

        // Normalize to ensure all rows have the same length
        self.normalize_data();
    }

    fn select_all(&mut self) {
        if !self.data.is_empty() {
            let max_cols = self.data.iter().map(|row| row.len()).max().unwrap_or(0);
            if max_cols > 0 {
                self.selection = Selection::CellRange {
                    start: (0, 0),
                    end: (self.data.len() - 1, max_cols - 1),
                };
                self.editing_cell = None;
            }
        }
    }
}

impl eframe::App for SpreadsheetApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Handle keyboard input - check shortcuts early before any UI
        let not_editing = self.editing_cell.is_none();

        // Handle high-level copy/paste/cut events from OS
        let mut do_copy = false;
        let mut do_paste = false;
        let mut do_cut = false;
        let mut paste_text: Option<String> = None;

        ctx.input(|i| {
            for event in &i.events {
                match event {
                    egui::Event::Copy => {
                        if not_editing {
                            do_copy = true;
                        }
                    }
                    egui::Event::Paste(text) => {
                        if not_editing {
                            do_paste = true;
                            paste_text = Some(text.clone());
                        }
                    }
                    egui::Event::Cut => {
                        if not_editing {
                            do_cut = true;
                        }
                    }
                    _ => {}
                }
            }
        });

        // Execute clipboard operations
        if do_copy {
            self.copy_selection();
        }
        if do_paste {
            if let Some(text) = paste_text {
                self.save_undo_state();
                self.paste_text(&text);
            }
        }
        if do_cut {
            self.cut_selection();
        }

        // Handle other keyboard shortcuts
        if not_editing && ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::A)) {
            self.select_all();
        }
        if ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::Y)) {
            self.redo();
        }
        if ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::Z)) {
            self.undo();
        }

        let mut start_editing_with: Option<String> = None;
        let mut move_selection: Option<(isize, isize)> = None; // (row_delta, col_delta)
        let mut extend_selection = false;
        let current_editing_cell = self.editing_cell;

        ctx.input(|i| {
            if i.key_pressed(egui::Key::Delete) || i.key_pressed(egui::Key::Backspace) {
                if self.editing_cell.is_none() {
                    self.save_undo_state();
                    self.clear_selection();
                }
            }
            if i.key_pressed(egui::Key::Escape) {
                self.selection = Selection::None;
                self.editing_cell = None;
            }
            // Clear drag state when mouse released
            if i.pointer.primary_released() {
                self.drag_start = None;
            }

            // Handle Enter when cell is being edited (arrow keys work normally for cursor movement)
            if current_editing_cell.is_some() {
                if i.key_pressed(egui::Key::Enter) {
                    move_selection = Some((1, 0)); // Move down
                }
            }

            // Handle arrow keys when cell is selected (not editing)
            if self.editing_cell.is_none() {
                extend_selection = i.modifiers.shift;

                if i.key_pressed(egui::Key::ArrowUp) {
                    move_selection = Some((-1, 0));
                } else if i.key_pressed(egui::Key::ArrowDown) {
                    move_selection = Some((1, 0));
                } else if i.key_pressed(egui::Key::ArrowLeft) {
                    move_selection = Some((0, -1));
                } else if i.key_pressed(egui::Key::ArrowRight) {
                    move_selection = Some((0, 1));
                }
            }

            // Start editing on text input when single cell is selected
            if self.editing_cell.is_none() {
                if let Selection::CellRange { start, end } = &self.selection {
                    if start == end {
                        // Single cell selected, check for text input
                        for event in &i.events {
                            if let egui::Event::Text(text) = event {
                                start_editing_with = Some(text.clone());
                                break;
                            }
                        }
                    }
                }
            }
        });

        // Start editing if text was typed
        if let Some(text) = start_editing_with {
            if let Selection::CellRange { start, end } = &self.selection {
                if start == end {
                    self.editing_cell = Some(*start);
                    self.edit_buffer = text;
                    self.selection = Selection::None;
                }
            }
        }

        egui::TopBottomPanel::top("menu_bar").show(ctx, |ui| {
            egui::MenuBar::new().ui(ui, |ui| {
                ui.menu_button("File", |ui| {
                    if ui.button("Open CSV").clicked() {
                        if let Some(path) = rfd::FileDialog::new()
                            .add_filter("CSV", &["csv"])
                            .pick_file()
                        {
                            self.load_csv(path);
                        }
                        ui.close();
                    }

                    if ui.button("Save").clicked() {
                        if let Some(ref path) = self.file_path {
                            if let Err(e) = self.save_csv(path) {
                                eprintln!("Error saving CSV: {}", e);
                            }
                        }
                        ui.close();
                    }

                    if ui.button("Save As...").clicked() {
                        if let Some(path) = rfd::FileDialog::new()
                            .add_filter("CSV", &["csv"])
                            .save_file()
                        {
                            if let Err(e) = self.save_csv(&path) {
                                eprintln!("Error saving CSV: {}", e);
                            } else {
                                self.file_path = Some(path);
                            }
                        }
                        ui.close();
                    }

                    if ui.button("New").clicked() {
                        self.data = vec![vec![String::new(); 10]; 20];
                        self.file_path = None;
                        ui.close();
                    }
                });

                ui.menu_button("Edit", |ui| {
                    if ui.button("Add Row").clicked() {
                        self.add_row();
                        ui.close();
                    }

                    if ui.button("Add Column").clicked() {
                        self.add_column();
                        ui.close();
                    }
                });
            });
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            let num_rows = self.data.len();
            let num_cols = self.data.iter().map(|r| r.len()).max().unwrap_or(0);
            let row_height = 25.0;

            // Clone selection for use in closures (before any updates)
            let current_selection = self.selection.clone();

            // Track pending operations
            let mut delete_row: Option<usize> = None;
            let mut delete_col: Option<usize> = None;
            let mut insert_row_at: Option<usize> = None;
            let mut insert_col_at: Option<usize> = None;
            let mut drag_end_cell: Option<(usize, usize)> = None;

            let mut table = TableBuilder::new(ui)
                .striped(true)
                .resizable(true)
                .cell_layout(egui::Layout::left_to_right(egui::Align::Center))
                .column(Column::initial(50.0).at_least(30.0)); // Row index column

            // Add data columns with custom widths
            for col_idx in 0..num_cols {
                let width = self.get_column_width(col_idx);
                table = table.column(Column::initial(width).at_least(30.0).resizable(true));
            }

            table
                .header(row_height, |mut header| {
                    // Corner cell
                    header.col(|_ui| {});

                    // Column headers
                    for col_idx in 0..num_cols {
                        header.col(|ui| {
                            let is_col_selected = matches!(&current_selection, Selection::Column(c) if *c == col_idx);

                            // Allocate space and draw background if selected
                            let (rect, response) = ui.allocate_exact_size(
                                ui.available_size(),
                                egui::Sense::click()
                            );

                            if is_col_selected {
                                ui.painter().rect_filled(rect, 0.0, egui::Color32::from_rgb(100, 150, 200));
                            }

                            // Draw column letter
                            ui.painter().text(
                                rect.center(),
                                egui::Align2::CENTER_CENTER,
                                Self::col_index_to_letter(col_idx),
                                egui::FontId::default(),
                                ui.visuals().text_color()
                            );

                            if response.clicked() {
                                self.selection = Selection::Column(col_idx);
                                self.editing_cell = None;
                            }

                            response.context_menu(|ui| {
                                if ui.button("Insert Column Left").clicked() {
                                    insert_col_at = Some(col_idx);
                                    ui.close();
                                }
                                if ui.button("Insert Column Right").clicked() {
                                    insert_col_at = Some(col_idx + 1);
                                    ui.close();
                                }
                                ui.separator();
                                if ui.button("Delete Column").clicked() {
                                    delete_col = Some(col_idx);
                                    ui.close();
                                }
                            });
                        });
                    }
                })
                .body(|body| {
                    body.rows(row_height, num_rows, |mut row| {
                        let row_idx = row.index();
                        let is_row_selected = matches!(&current_selection, Selection::Row(r) if *r == row_idx);

                        // Row number
                        row.col(|ui| {
                            // Allocate space and draw background if selected
                            let (rect, response) = ui.allocate_exact_size(
                                ui.available_size(),
                                egui::Sense::click()
                            );

                            if is_row_selected {
                                ui.painter().rect_filled(rect, 0.0, egui::Color32::from_rgb(100, 150, 200));
                            }

                            // Draw row number
                            ui.painter().text(
                                rect.center(),
                                egui::Align2::CENTER_CENTER,
                                (row_idx + 1).to_string(),
                                egui::FontId::default(),
                                ui.visuals().text_color()
                            );

                            if response.clicked() {
                                self.selection = Selection::Row(row_idx);
                                self.editing_cell = None;
                            }

                            response.context_menu(|ui| {
                                if ui.button("Insert Row Above").clicked() {
                                    insert_row_at = Some(row_idx);
                                    ui.close();
                                }
                                if ui.button("Insert Row Below").clicked() {
                                    insert_row_at = Some(row_idx + 1);
                                    ui.close();
                                }
                                ui.separator();
                                if ui.button("Delete Row").clicked() {
                                    delete_row = Some(row_idx);
                                    ui.close();
                                }
                            });
                        });

                        // Data cells
                        for col_idx in 0..num_cols {
                            row.col(|ui| {
                                let cell_id = (row_idx, col_idx);
                                let is_editing = self.editing_cell == Some(cell_id);

                                // Calculate is_selected inline without calling self methods
                                let is_selected = match &current_selection {
                                    Selection::None => false,
                                    Selection::CellRange { start, end } => {
                                        let (r1, c1) = *start;
                                        let (r2, c2) = *end;
                                        let (min_r, max_r) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
                                        let (min_c, max_c) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
                                        row_idx >= min_r && row_idx <= max_r && col_idx >= min_c && col_idx <= max_c
                                    }
                                    Selection::Column(c) => col_idx == *c,
                                    Selection::Row(r) => row_idx == *r,
                                };

                                if let Some(row_data) = self.data.get_mut(row_idx) {
                                    if col_idx >= row_data.len() {
                                        return; // Skip if column doesn't exist yet
                                    }
                                    let cell_val = &mut row_data[col_idx];
                                        if is_editing {
                                            let text_edit = egui::TextEdit::singleline(&mut self.edit_buffer)
                                                .frame(true);

                                            let response = ui.add(text_edit);

                                            // Check if Enter was pressed to move down
                                            let enter_pressed = ui.input(|i| i.key_pressed(egui::Key::Enter));

                                            if response.lost_focus() || enter_pressed {
                                                *cell_val = self.edit_buffer.clone();
                                                self.editing_cell = None;
                                            }

                                            response.request_focus();
                                        } else {
                                            // Create an interactive area that fills the cell
                                            let (rect, response) = ui.allocate_exact_size(
                                                ui.available_size(),
                                                egui::Sense::click_and_drag()
                                            );

                                            // Draw selection background
                                            if is_selected {
                                                ui.painter().rect_filled(
                                                    rect,
                                                    0.0,
                                                    egui::Color32::from_rgb(180, 210, 240)
                                                );
                                            }

                                            // Draw the text
                                            let text_pos = rect.left_center() + egui::vec2(4.0, 0.0);
                                            ui.painter().text(
                                                text_pos,
                                                egui::Align2::LEFT_CENTER,
                                                &*cell_val,
                                                egui::FontId::default(),
                                                ui.visuals().text_color()
                                            );

                                            // Double-click to edit
                                            if response.double_clicked() {
                                                self.editing_cell = Some(cell_id);
                                                self.edit_buffer = cell_val.clone();
                                                self.selection = Selection::None;
                                                self.drag_start = None;
                                            }
                                            // Start drag selection
                                            else if response.is_pointer_button_down_on() {
                                                self.drag_start = Some(cell_id);
                                                self.selection = Selection::CellRange { start: cell_id, end: cell_id };
                                                self.editing_cell = None;
                                            }

                                            // Track drag end cell for later update (avoid flicker)
                                            if self.drag_start.is_some() && ui.input(|i| i.pointer.primary_down()) {
                                                if let Some(pos) = ui.input(|i| i.pointer.hover_pos()) {
                                                    // Expand rect to handle edge cases (dividers, fast dragging)
                                                    let expanded_rect = rect.expand(5.0);
                                                    if expanded_rect.contains(pos) {
                                                        drag_end_cell = Some(cell_id);
                                                    }
                                                }
                                            }

                                            response.context_menu(|ui| {
                                                if ui.button("Clear").clicked() {
                                                    cell_val.clear();
                                                    ui.close();
                                                }
                                            });
                                        }
                                }
                            });
                        }
                    });
                });

            // Update selection based on drag AFTER table render
            if let Some(end_cell) = drag_end_cell {
                if let Some(start) = self.drag_start {
                    self.selection = Selection::CellRange { start, end: end_cell };
                    // Request immediate repaint to show selection update without flicker
                    ctx.request_repaint();
                }
            }

            // Request continuous repaints while dragging for smooth selection updates
            if self.drag_start.is_some() && ui.input(|i| i.pointer.primary_down()) {
                ctx.request_repaint();
            }

            // Process pending operations after UI rendering
            if let Some(col_idx) = insert_col_at {
                self.insert_column_at(col_idx);
            }
            if let Some(row_idx) = insert_row_at {
                self.insert_row_at(row_idx);
            }
            if let Some(col_idx) = delete_col {
                self.delete_column(col_idx);
            }
            if let Some(row_idx) = delete_row {
                self.delete_row(row_idx);
            }

            // Clear drag state when mouse released
            if ui.input(|i| i.pointer.primary_released()) {
                self.drag_start = None;
            }

            // Handle cell navigation (Arrow keys/Enter)
            if let Some((row_delta, col_delta)) = move_selection {
                self.editing_cell = None;

                // Get current position and selection anchor
                let (anchor, current_pos) = if let Some((row, col)) = current_editing_cell {
                    ((row, col), (row, col))
                } else if let Selection::CellRange { start, end } = &self.selection {
                    (*start, *end)
                } else {
                    ((0, 0), (0, 0))
                };

                let new_row = (current_pos.0 as isize + row_delta).max(0).min((num_rows - 1) as isize) as usize;
                let new_col = (current_pos.1 as isize + col_delta).max(0).min((num_cols - 1) as isize) as usize;

                if extend_selection {
                    // Extend selection from anchor to new position
                    self.selection = Selection::CellRange {
                        start: anchor,
                        end: (new_row, new_col)
                    };
                } else {
                    // Move to new cell
                    self.selection = Selection::CellRange {
                        start: (new_row, new_col),
                        end: (new_row, new_col)
                    };
                }
            }
        });
    }
}
