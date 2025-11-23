use eframe::egui;
use egui_extras::{TableBuilder, Column};
use std::collections::HashMap;
use std::path::PathBuf;

// WASM-specific imports
#[cfg(target_arch = "wasm32")]
use std::sync::{Arc, Mutex};

// Structure to hold data loaded asynchronously (WASM)
#[cfg(target_arch = "wasm32")]
#[derive(Default, Clone)]
struct AsyncFileResult {
    data: Option<Vec<u8>>,
    filename: Option<String>,
}

// Platform-specific clipboard implementation
#[cfg(not(target_arch = "wasm32"))]
struct ClipboardContext {
    clipboard: arboard::Clipboard,
}

#[cfg(target_arch = "wasm32")]
struct ClipboardContext {}

impl ClipboardContext {
    #[cfg(not(target_arch = "wasm32"))]
    fn new() -> Option<Self> {
        arboard::Clipboard::new().ok().map(|clipboard| Self { clipboard })
    }

    #[cfg(target_arch = "wasm32")]
    fn new() -> Option<Self> {
        Some(Self {})
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn set_text(&mut self, text: String) -> Result<(), Box<dyn std::error::Error>> {
        self.clipboard.set_text(text)?;
        Ok(())
    }

    #[cfg(target_arch = "wasm32")]
    fn set_text(&mut self, text: String) -> Result<(), Box<dyn std::error::Error>> {
        if let Some(window) = web_sys::window() {
            let navigator = window.navigator();
            let clipboard = navigator.clipboard();
            let _ = clipboard.write_text(&text);
        }
        Ok(())
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn get_text(&mut self) -> Result<String, Box<dyn std::error::Error>> {
        Ok(self.clipboard.get_text()?)
    }

    #[cfg(target_arch = "wasm32")]
    fn get_text(&mut self) -> Result<String, Box<dyn std::error::Error>> {
        // In WASM, we'll rely on egui's paste events instead
        Err("Use egui paste events in WASM".into())
    }
}

#[derive(Debug, Clone, PartialEq)]
enum Selection {
    None,
    CellRange { start: (usize, usize), end: (usize, usize) },
    Column(usize),
    Row(usize),
}

#[derive(Debug, Clone, PartialEq)]
enum PendingAction {
    None,
    NewFile,
    OpenFile,
    Exit,
}

#[cfg(not(target_arch = "wasm32"))]
fn load_icon() -> Option<egui::IconData> {
    let icon_bytes = include_bytes!("../logo-nobg.png");
    let image = image::load_from_memory(icon_bytes).ok()?;
    let rgba = image.to_rgba8();
    let (width, height) = rgba.dimensions();

    Some(egui::IconData {
        rgba: rgba.into_raw(),
        width: width as u32,
        height: height as u32,
    })
}

// Native entry point
#[cfg(not(target_arch = "wasm32"))]
fn main() -> eframe::Result<()> {
    let mut viewport = egui::ViewportBuilder::default()
        .with_inner_size([1200.0, 800.0]);

    if let Some(icon) = load_icon() {
        viewport = viewport.with_icon(icon);
    }

    let options = eframe::NativeOptions {
        viewport,
        ..Default::default()
    };

    eframe::run_native(
        "GridView",
        options,
        Box::new(|_cc| Ok(Box::new(SpreadsheetApp::default()))),
    )
}

// WASM entry point
#[cfg(target_arch = "wasm32")]
fn main() {
    use wasm_bindgen::JsCast;

    // Redirect tracing to console.log and friends:
    eframe::WebLogger::init(log::LevelFilter::Debug).ok();

    let web_options = eframe::WebOptions::default();

    wasm_bindgen_futures::spawn_local(async {
        // Get the canvas element
        let document = web_sys::window()
            .expect("No window")
            .document()
            .expect("No document");

        let canvas = document
            .get_element_by_id("the_canvas_id")
            .expect("Canvas not found")
            .dyn_into::<web_sys::HtmlCanvasElement>()
            .expect("Element is not a canvas");

        let start_result = eframe::WebRunner::new()
            .start(
                canvas,
                web_options,
                Box::new(|_cc| Ok(Box::new(SpreadsheetApp::default()))),
            )
            .await;

        // Remove the loading text and spinner:
        let loading_text = web_sys::window()
            .and_then(|w| w.document())
            .and_then(|d| d.get_element_by_id("loading_text"));
        if let Some(loading_text) = loading_text {
            match start_result {
                Ok(_) => {
                    loading_text.remove();
                }
                Err(e) => {
                    loading_text.set_inner_html(
                        "<p> The app has crashed. See the developer console for details. </p>",
                    );
                    panic!("Failed to start eframe: {e:?}");
                }
            }
        }
    });
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
    clipboard: ClipboardContext,
    undo_stack: Vec<Vec<Vec<String>>>,
    redo_stack: Vec<Vec<Vec<String>>>,
    pending_action: PendingAction,
    has_unsaved_changes: bool,
    allowed_to_close: bool,
    table_id_salt: u64, // Change this to reset table state
    dark_mode: bool,
    // Search functionality
    search_window_open: bool,
    search_query: String,
    search_case_sensitive: bool,
    search_results: Vec<(usize, usize)>,
    current_search_result: usize,
    // Sort tracking
    sorted_column: Option<usize>,
    sort_ascending: bool,
    freeze_top_row: bool,
    #[cfg(target_arch = "wasm32")]
    async_file_loading: Arc<Mutex<AsyncFileResult>>,
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
            clipboard: ClipboardContext::new().unwrap(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            pending_action: PendingAction::None,
            has_unsaved_changes: false,
            allowed_to_close: false,
            table_id_salt: 0,
            dark_mode: true, // Default to dark mode
            search_window_open: false,
            search_query: String::new(),
            search_case_sensitive: false,
            search_results: Vec::new(),
            current_search_result: 0,
            sorted_column: None,
            sort_ascending: true,
            freeze_top_row: true, // Default to frozen, common for CSVs
            #[cfg(target_arch = "wasm32")]
            async_file_loading: Arc::new(Mutex::new(AsyncFileResult::default())),
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

    #[allow(dead_code)]
    fn load_csv(&mut self, path: PathBuf) {
        #[cfg(not(target_arch = "wasm32"))]
        {
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
                    self.has_unsaved_changes = false;
                }
                Err(e) => {
                    eprintln!("Error loading CSV: {}", e);
                }
            }
        }
        #[cfg(target_arch = "wasm32")]
        {
            // Store path for display purposes
            self.file_path = Some(path);
        }
    }

    #[allow(dead_code)]
    fn load_csv_from_bytes(&mut self, bytes: &[u8], filename: String) {
        let mut reader = csv::Reader::from_reader(bytes);
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
        self.file_path = Some(PathBuf::from(filename));
        self.has_unsaved_changes = false;
    }

    fn save_csv(&self, path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
        #[cfg(not(target_arch = "wasm32"))]
        {
            let mut writer = csv::Writer::from_path(path)?;

            for row in &self.data {
                writer.write_record(row)?;
            }

            writer.flush()?;
        }
        #[cfg(target_arch = "wasm32")]
        {
            // WASM save will be handled via download
            let _ = path; // Suppress unused warning
        }
        Ok(())
    }

    #[allow(dead_code)]
    fn save_csv_to_bytes(&self) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut writer = csv::Writer::from_writer(Vec::new());

        for row in &self.data {
            writer.write_record(row)?;
        }

        writer.flush()?;
        Ok(writer.into_inner()?)
    }

    fn add_row(&mut self) {
        let cols = self.data.first().map(|r| r.len()).unwrap_or(10);
        self.data.push(vec![String::new(); cols]);
        self.has_unsaved_changes = true;
    }

    fn add_column(&mut self) {
        if self.data.is_empty() {
            self.data.push(vec![String::new()]);
        } else {
            for row in &mut self.data {
                row.push(String::new());
            }
        }
        self.has_unsaved_changes = true;
    }

    fn insert_row_at(&mut self, row_idx: usize) {
        let cols = self.data.first().map(|r| r.len()).unwrap_or(10);
        self.data.insert(row_idx, vec![String::new(); cols]);
        self.has_unsaved_changes = true;

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
        self.has_unsaved_changes = true;

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
            self.has_unsaved_changes = true;
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
        self.has_unsaved_changes = true;
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
        self.has_unsaved_changes = true;
        // Limit undo stack to 50 entries
        if self.undo_stack.len() > 50 {
            self.undo_stack.remove(0);
        }
    }

    fn undo(&mut self) {
        if let Some(prev_state) = self.undo_stack.pop() {
            self.redo_stack.push(self.data.clone());
            self.data = prev_state;
            self.has_unsaved_changes = true;
            // Clear sort indicator since data state changed
            self.sorted_column = None;
        }
    }

    fn redo(&mut self) {
        if let Some(next_state) = self.redo_stack.pop() {
            self.undo_stack.push(self.data.clone());
            self.data = next_state;
            self.has_unsaved_changes = true;
            // Clear sort indicator since data state changed
            self.sorted_column = None;
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

    fn perform_search(&mut self) {
        self.search_results.clear();
        self.current_search_result = 0;

        if self.search_query.is_empty() {
            return;
        }

        let query = if self.search_case_sensitive {
            self.search_query.clone()
        } else {
            self.search_query.to_lowercase()
        };

        for (row_idx, row) in self.data.iter().enumerate() {
            for (col_idx, cell) in row.iter().enumerate() {
                let cell_text = if self.search_case_sensitive {
                    cell.clone()
                } else {
                    cell.to_lowercase()
                };

                if cell_text.contains(&query) {
                    self.search_results.push((row_idx, col_idx));
                }
            }
        }
    }

    fn go_to_next_search_result(&mut self) {
        if !self.search_results.is_empty() {
            self.current_search_result = (self.current_search_result + 1) % self.search_results.len();
            let (row, col) = self.search_results[self.current_search_result];
            self.selection = Selection::CellRange {
                start: (row, col),
                end: (row, col),
            };
            self.editing_cell = None;
        }
    }

    fn go_to_prev_search_result(&mut self) {
        if !self.search_results.is_empty() {
            if self.current_search_result == 0 {
                self.current_search_result = self.search_results.len() - 1;
            } else {
                self.current_search_result -= 1;
            }
            let (row, col) = self.search_results[self.current_search_result];
            self.selection = Selection::CellRange {
                start: (row, col),
                end: (row, col),
            };
            self.editing_cell = None;
        }
    }

    fn sort_by_column(&mut self, col_idx: usize, ascending: bool) {
        if self.data.is_empty() {
            return;
        }

        self.save_undo_state();

        if self.freeze_top_row && self.data.len() > 1 {
            // Sort only data rows (skip first row which is the header)
            let mut header = self.data[0].clone();
            let mut data_rows: Vec<Vec<String>> = self.data[1..].to_vec();

            data_rows.sort_by(|a, b| {
                let a_val = a.get(col_idx).map(|s| s.as_str()).unwrap_or("");
                let b_val = b.get(col_idx).map(|s| s.as_str()).unwrap_or("");

                // Try to parse as numbers first
                let cmp = match (a_val.parse::<f64>(), b_val.parse::<f64>()) {
                    (Ok(a_num), Ok(b_num)) => a_num.partial_cmp(&b_num).unwrap_or(std::cmp::Ordering::Equal),
                    _ => a_val.cmp(b_val),
                };

                if ascending {
                    cmp
                } else {
                    cmp.reverse()
                }
            });

            // Reconstruct with header at top
            self.data.clear();
            self.data.push(header);
            self.data.extend(data_rows);
        } else {
            // Sort all data (no frozen row)
            self.data.sort_by(|a, b| {
                let a_val = a.get(col_idx).map(|s| s.as_str()).unwrap_or("");
                let b_val = b.get(col_idx).map(|s| s.as_str()).unwrap_or("");

                // Try to parse as numbers first
                let cmp = match (a_val.parse::<f64>(), b_val.parse::<f64>()) {
                    (Ok(a_num), Ok(b_num)) => a_num.partial_cmp(&b_num).unwrap_or(std::cmp::Ordering::Equal),
                    _ => a_val.cmp(b_val),
                };

                if ascending {
                    cmp
                } else {
                    cmp.reverse()
                }
            });
        }

        // Track which column is sorted
        self.sorted_column = Some(col_idx);
        self.sort_ascending = ascending;
        self.has_unsaved_changes = true;
    }

    #[cfg(target_arch = "wasm32")]
    fn download_file(&self, data: &[u8], filename: &str) {
        use wasm_bindgen::JsCast;

        if let Some(window) = web_sys::window() {
            if let Some(document) = window.document() {
                // Create a blob from the data
                let array = js_sys::Uint8Array::new_with_length(data.len() as u32);
                array.copy_from(data);
                let blob_parts = js_sys::Array::new();
                blob_parts.push(&array);

                if let Ok(blob) = web_sys::Blob::new_with_u8_array_sequence(&blob_parts) {
                    if let Ok(url) = web_sys::Url::create_object_url_with_blob(&blob) {
                        // Create download link
                        if let Ok(element) = document.create_element("a") {
                            if let Ok(anchor) = element.dyn_into::<web_sys::HtmlAnchorElement>() {
                                anchor.set_href(&url);
                                anchor.set_download(filename);
                                anchor.click();
                                let _ = web_sys::Url::revoke_object_url(&url);
                            }
                        }
                    }
                }
            }
        }
    }

    fn trigger_open_file(&mut self) {
        #[cfg(not(target_arch = "wasm32"))]
        {
            if let Some(path) = rfd::FileDialog::new()
                .add_filter("CSV", &["csv"])
                .pick_file()
            {
                self.load_csv(path);
            }
        }

        #[cfg(target_arch = "wasm32")]
        {
            let result_clone = self.async_file_loading.clone();

            // Spawn a future to handle the file picker
            wasm_bindgen_futures::spawn_local(async move {
                // rfd::AsyncFileDialog works perfectly in WASM
                let file = rfd::AsyncFileDialog::new()
                    .add_filter("CSV", &["csv"])
                    .pick_file()
                    .await;

                if let Some(file_handle) = file {
                    let name = file_handle.file_name();
                    let bytes = file_handle.read().await; // Reads into Vec<u8>

                    // Lock the mutex and store the data
                    if let Ok(mut guard) = result_clone.lock() {
                        guard.data = Some(bytes);
                        guard.filename = Some(name);
                    }
                }
            });
        }
    }
}

impl eframe::App for SpreadsheetApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Poll for loaded files (WASM)
        #[cfg(target_arch = "wasm32")]
        {
            let mut loaded_data = None;

            // Minimal lock scope
            if let Ok(mut guard) = self.async_file_loading.lock() {
                if guard.data.is_some() {
                    // Move data out of the mutex
                    loaded_data = Some((
                        guard.data.take().unwrap(),
                        guard.filename.take().unwrap(),
                    ));
                }
            }

            if let Some((bytes, filename)) = loaded_data {
                self.load_csv_from_bytes(&bytes, filename);
            }
        }

        // Handle drag-and-drop files
        ctx.input(|i| {
            if !i.raw.dropped_files.is_empty() {
                for file in &i.raw.dropped_files {
                    // Native: use path directly
                    #[cfg(not(target_arch = "wasm32"))]
                    {
                        if let Some(path) = &file.path {
                            // Only load CSV files
                            if path.extension().and_then(|s| s.to_str()) == Some("csv") {
                                self.load_csv(path.clone());
                            }
                        }
                    }

                    // WASM: use bytes from dropped file
                    #[cfg(target_arch = "wasm32")]
                    {
                        if let Some(bytes) = &file.bytes {
                            let filename = file.name.clone();
                            // Only load CSV files
                            if filename.ends_with(".csv") {
                                self.load_csv_from_bytes(bytes, filename);
                            }
                        }
                    }
                }
            }
        });

        // Intercept window close button (X)
        if ctx.input(|i| i.viewport().close_requested()) {
            if self.allowed_to_close {
                // User confirmed exit - allow window to close
            } else if self.has_unsaved_changes {
                // Prevent close and show confirmation modal
                ctx.send_viewport_cmd(egui::ViewportCommand::CancelClose);
                self.pending_action = PendingAction::Exit;
            }
            // If no unsaved changes, allow window to close
        }

        // Update window title to show filename and unsaved changes indicator
        let title = if let Some(ref path) = self.file_path {
            let filename = path.file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("Untitled");
            if self.has_unsaved_changes {
                format!("GridView - {} *", filename)
            } else {
                format!("GridView - {}", filename)
            }
        } else {
            if self.has_unsaved_changes {
                "GridView *".to_string()
            } else {
                "GridView".to_string()
            }
        };
        ctx.send_viewport_cmd(egui::ViewportCommand::Title(title));

        // Apply theme
        if self.dark_mode {
            let mut visuals = egui::Visuals::dark();
            // Make text whiter for better contrast
            visuals.override_text_color = Some(egui::Color32::from_rgb(240, 240, 240));
            ctx.set_visuals(visuals);
        } else {
            ctx.set_visuals(egui::Visuals::light());
        }

        // Handle keyboard input - check shortcuts early before any UI
        let not_editing = self.editing_cell.is_none();

        // File operation shortcuts (Cmd/Ctrl + S/N/O/Shift+S)
        if not_editing {
            // Cmd+N - New File
            if ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::N)) {
                if self.has_unsaved_changes {
                    self.pending_action = PendingAction::NewFile;
                } else {
                    self.data = vec![vec![String::new(); 10]; 20];
                    self.file_path = None;
                }
            }

            // Cmd+O - Open File
            if ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::O)) {
                if self.has_unsaved_changes {
                    self.pending_action = PendingAction::OpenFile;
                } else {
                    #[cfg(not(target_arch = "wasm32"))]
                    {
                        if let Some(path) = rfd::FileDialog::new()
                            .add_filter("CSV", &["csv"])
                            .pick_file()
                        {
                            self.load_csv(path);
                        }
                    }
                    #[cfg(target_arch = "wasm32")]
                    {
                        self.trigger_open_file();
                    }
                }
            }

            // Cmd+Shift+S - Save As
            if ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND.plus(egui::Modifiers::SHIFT), egui::Key::S)) {
                #[cfg(not(target_arch = "wasm32"))]
                {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("CSV", &["csv"])
                        .save_file()
                    {
                        if let Err(e) = self.save_csv(&path) {
                            eprintln!("Error saving CSV: {}", e);
                        } else {
                            self.file_path = Some(path);
                            self.has_unsaved_changes = false;
                        }
                    }
                }
                #[cfg(target_arch = "wasm32")]
                {
                    // WASM: Trigger download
                    if let Ok(bytes) = self.save_csv_to_bytes() {
                        self.download_file(&bytes, "spreadsheet.csv");
                        self.has_unsaved_changes = false;
                    }
                }
            }
            // Cmd+S - Save (must come after Cmd+Shift+S check)
            else if ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::S)) {
                if let Some(ref path) = self.file_path {
                    if let Err(e) = self.save_csv(path) {
                        eprintln!("Error saving CSV: {}", e);
                    } else {
                        self.has_unsaved_changes = false;
                    }
                }
            }
        }

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
        if not_editing && !self.search_window_open && ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::A)) {
            self.select_all();
        }
        if not_editing && ctx.input_mut(|i| i.consume_key(egui::Modifiers::COMMAND, egui::Key::F)) {
            self.search_window_open = true;
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

            // Start editing on text input when single cell is selected (but not when search window is open)
            if self.editing_cell.is_none() && !self.search_window_open {
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
                    if ui.button("New").clicked() {
                        if self.has_unsaved_changes {
                            self.pending_action = PendingAction::NewFile;
                        } else {
                            // No unsaved changes, create new file directly
                            self.data = vec![vec![String::new(); 10]; 20];
                            self.file_path = None;
                        }
                        ui.close();
                    }

                    if ui.button("Open CSV").clicked() {
                        if self.has_unsaved_changes {
                            self.pending_action = PendingAction::OpenFile;
                        } else {
                            // No unsaved changes, open file directly
                            #[cfg(not(target_arch = "wasm32"))]
                            {
                                if let Some(path) = rfd::FileDialog::new()
                                    .add_filter("CSV", &["csv"])
                                    .pick_file()
                                {
                                    self.load_csv(path);
                                }
                            }
                            #[cfg(target_arch = "wasm32")]
                            {
                                self.trigger_open_file();
                            }
                        }
                        ui.close();
                    }

                    if ui.button("Save").clicked() {
                        if let Some(ref path) = self.file_path {
                            if let Err(e) = self.save_csv(path) {
                                eprintln!("Error saving CSV: {}", e);
                            } else {
                                self.has_unsaved_changes = false;
                            }
                        }
                        ui.close();
                    }

                    if ui.button("Save As...").clicked() {
                        #[cfg(not(target_arch = "wasm32"))]
                        {
                            if let Some(path) = rfd::FileDialog::new()
                                .add_filter("CSV", &["csv"])
                                .save_file()
                            {
                                if let Err(e) = self.save_csv(&path) {
                                    eprintln!("Error saving CSV: {}", e);
                                } else {
                                    self.file_path = Some(path);
                                    self.has_unsaved_changes = false;
                                }
                            }
                        }
                        #[cfg(target_arch = "wasm32")]
                        {
                            if let Ok(bytes) = self.save_csv_to_bytes() {
                                self.download_file(&bytes, "spreadsheet.csv");
                                self.has_unsaved_changes = false;
                            }
                        }
                        ui.close();
                    }
                });

                ui.menu_button("Edit", |ui| {
                    if ui.button("Cut").clicked() {
                        self.cut_selection();
                        ui.close();
                    }

                    if ui.button("Copy").clicked() {
                        self.copy_selection();
                        ui.close();
                    }

                    if ui.button("Paste").clicked() {
                        if let Ok(text) = self.clipboard.get_text() {
                            self.save_undo_state();
                            self.paste_text(&text);
                        }
                        ui.close();
                    }

                    ui.separator();

                    if ui.button("Add Row").clicked() {
                        self.add_row();
                        ui.close();
                    }

                    if ui.button("Add Column").clicked() {
                        self.add_column();
                        ui.close();
                    }
                });

                ui.menu_button("View", |ui| {
                    if ui.checkbox(&mut self.freeze_top_row, "Freeze Top Row").clicked() {
                        ui.close();
                    }

                    ui.separator();

                    if ui.button("Reset Column Widths").clicked() {
                        self.column_widths.clear();
                        self.table_id_salt += 1; // Change table ID to reset egui's internal state
                        ui.close();
                    }

                    ui.separator();

                    let theme_label = if self.dark_mode { "Light Mode" } else { "Dark Mode" };
                    if ui.button(theme_label).clicked() {
                        self.dark_mode = !self.dark_mode;
                        ui.close();
                    }
                });
            });
        });

        // Always render the central panel, but disable interaction when modal is open
        egui::CentralPanel::default().show(ctx, |ui| {
            let num_rows = self.data.len();
            let num_cols = self.data.iter().map(|r| r.len()).max().unwrap_or(0);
            let row_height = 25.0;

            // Wrap everything in add_enabled_ui to disable interaction when modal is open
            ui.add_enabled_ui(self.pending_action == PendingAction::None, |ui| {
                // Wrap in ScrollArea for horizontal scrolling
                egui::ScrollArea::both()
                    .auto_shrink([false; 2])
                    .show(ui, |ui| {
                        // Remove spacing between cells to create continuous grid
                        ui.style_mut().spacing.item_spacing = egui::vec2(0.0, 0.0);

                        // Track if we should save current edit when clicking away
                        let mut save_current_edit = false;
                        let previous_editing_cell = self.editing_cell; // Capture BEFORE we change it

                        // Clone selection for use in closures (before any updates)
                        let current_selection = self.selection.clone();

                        // Track pending operations
                        let mut delete_row: Option<usize> = None;
                        let mut delete_col: Option<usize> = None;
                        let mut insert_row_at: Option<usize> = None;
                        let mut insert_col_at: Option<usize> = None;
                        let mut drag_end_cell: Option<(usize, usize)> = None;
                        let mut clear_cell: Option<(usize, usize)> = None;

                        let mut table = TableBuilder::new(ui)
                .id_salt(self.table_id_salt) // Use salt to reset table state
                .striped(true)
                .resizable(true)
                .cell_layout(egui::Layout::left_to_right(egui::Align::Center))
                .vscroll(false) // Disable internal scroll since we have outer ScrollArea
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

                            // Draw column letter with sort indicator
                            let mut col_text = Self::col_index_to_letter(col_idx);
                            if self.sorted_column == Some(col_idx) {
                                col_text.push(' ');
                                col_text.push(if self.sort_ascending { '^' } else { 'v' });
                            }
                            ui.painter().text(
                                rect.center(),
                                egui::Align2::CENTER_CENTER,
                                col_text,
                                egui::FontId::default(),
                                ui.visuals().text_color()
                            );

                            if response.clicked() {
                                self.selection = Selection::Column(col_idx);
                                self.editing_cell = None;
                            }

                            response.context_menu(|ui| {
                                if ui.button("Sort Ascending").clicked() {
                                    self.sort_by_column(col_idx, true);
                                    ui.close();
                                }
                                if ui.button("Sort Descending").clicked() {
                                    self.sort_by_column(col_idx, false);
                                    ui.close();
                                }
                                ui.separator();
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

                                // Check if cell is in search results
                                let is_search_match = self.search_results.contains(&(row_idx, col_idx));
                                let is_current_search_result = !self.search_results.is_empty()
                                    && self.search_results.get(self.current_search_result) == Some(&(row_idx, col_idx));

                                if let Some(row_data) = self.data.get_mut(row_idx) {
                                    if col_idx >= row_data.len() {
                                        return; // Skip if column doesn't exist yet
                                    }
                                    let cell_val = &mut row_data[col_idx];
                                        // Create an interactive area that fills the cell
                                        let (rect, response) = ui.allocate_exact_size(
                                            ui.available_size(),
                                            egui::Sense::click_and_drag()
                                        );

                                        // Check if this is the frozen header row
                                        let is_frozen_header = self.freeze_top_row && row_idx == 0;

                                        // Draw cell background with priority: frozen header > current search result > search match > selected > normal
                                        // Use different colors for dark vs light mode
                                        let bg_color = if is_frozen_header {
                                            if self.dark_mode {
                                                egui::Color32::from_rgb(50, 50, 60) // Slightly lighter than background for dark mode
                                            } else {
                                                egui::Color32::from_rgb(230, 230, 240) // Slightly darker than background for light mode
                                            }
                                        } else if is_current_search_result {
                                            if self.dark_mode {
                                                egui::Color32::from_rgb(180, 100, 0) // Dark orange for dark mode
                                            } else {
                                                egui::Color32::from_rgb(255, 200, 100) // Light orange for light mode
                                            }
                                        } else if is_search_match {
                                            if self.dark_mode {
                                                egui::Color32::from_rgb(120, 100, 0) // Dark gold for dark mode
                                            } else {
                                                egui::Color32::from_rgb(255, 255, 150) // Light yellow for light mode
                                            }
                                        } else if is_selected {
                                            if self.dark_mode {
                                                egui::Color32::from_rgb(60, 90, 120) // Dark blue for dark mode
                                            } else {
                                                egui::Color32::from_rgb(180, 210, 240) // Light blue for light mode
                                            }
                                        } else {
                                            egui::Color32::TRANSPARENT
                                        };

                                        if bg_color != egui::Color32::TRANSPARENT {
                                            ui.painter().rect_filled(rect, 0.0, bg_color);
                                        }

                                        // Draw cell border (blue if editing, normal grid color otherwise)
                                        let border_color = if is_editing {
                                            egui::Color32::from_rgb(66, 133, 244) // Blue border when editing
                                        } else {
                                            ui.visuals().widgets.noninteractive.bg_stroke.color
                                        };

                                        let border_width = if is_editing { 2.0 } else { 0.5 };

                                        ui.painter().rect_stroke(
                                            rect,
                                            0.0,
                                            egui::Stroke::new(border_width, border_color),
                                            egui::epaint::StrokeKind::Inside
                                        );

                                        if is_editing {
                                            // Show text edit without frame, just cursor
                                            let edit_rect = rect.shrink2(egui::vec2(4.0, 2.0));
                                            let mut child_ui = ui.new_child(
                                                egui::UiBuilder::new()
                                                    .max_rect(edit_rect)
                                                    .layout(egui::Layout::left_to_right(egui::Align::Center))
                                            );

                                            let text_edit = egui::TextEdit::singleline(&mut self.edit_buffer)
                                                .frame(false);

                                            let edit_response = child_ui.add(text_edit);

                                            // Check if Enter was pressed to move down
                                            let enter_pressed = ui.input(|i| i.key_pressed(egui::Key::Enter));

                                            if edit_response.lost_focus() || enter_pressed {
                                                *cell_val = self.edit_buffer.clone();
                                                self.has_unsaved_changes = true;
                                                self.editing_cell = None;
                                            }

                                            edit_response.request_focus();
                                        } else {
                                            // Draw the text with clipping to prevent overflow
                                            let text_rect = rect.shrink2(egui::vec2(4.0, 0.0));
                                            let text_pos = rect.left_center() + egui::vec2(4.0, 0.0);

                                            // Use bold font for frozen header row
                                            let font_id = if is_frozen_header {
                                                egui::FontId::proportional(14.0)
                                            } else {
                                                egui::FontId::default()
                                            };

                                            ui.painter().with_clip_rect(text_rect).text(
                                                text_pos,
                                                egui::Align2::LEFT_CENTER,
                                                &*cell_val,
                                                font_id,
                                                ui.visuals().text_color()
                                            );

                                            // Double-click to edit
                                            if response.double_clicked() {
                                                save_current_edit = true;
                                                self.editing_cell = Some(cell_id);
                                                self.edit_buffer = cell_val.clone();
                                                self.selection = Selection::None;
                                                self.drag_start = None;
                                            }
                                            // Start drag selection
                                            else if response.is_pointer_button_down_on() {
                                                save_current_edit = true;
                                                self.drag_start = Some(cell_id);
                                                self.selection = Selection::CellRange { start: cell_id, end: cell_id };
                                                self.editing_cell = None;
                                            }

                                            // Track drag end cell for later update (avoid flicker)
                                            if self.drag_start.is_some() && ui.input(|i| i.pointer.primary_down()) {
                                                if let Some(pos) = ui.input(|i| i.pointer.hover_pos()) {
                                                    if rect.contains(pos) {
                                                        drag_end_cell = Some(cell_id);
                                                    }
                                                }
                                            }

                                            // Clear drag state when opening context menu (right-click)
                                            if response.secondary_clicked() {
                                                self.drag_start = None;
                                            }

                                            response.context_menu(|ui| {
                                                if ui.button("Cut").clicked() {
                                                    self.cut_selection();
                                                    ui.close();
                                                }
                                                if ui.button("Copy").clicked() {
                                                    self.copy_selection();
                                                    ui.close();
                                                }
                                                if ui.button("Paste").clicked() {
                                                    if let Ok(text) = self.clipboard.get_text() {
                                                        self.save_undo_state();
                                                        self.paste_text(&text);
                                                    }
                                                    ui.close();
                                                }
                                                ui.separator();
                                                if ui.button("Clear").clicked() {
                                                    clear_cell = Some(cell_id);
                                                    ui.close();
                                                }
                                            });
                                        }
                                }
                            });
                        }
                    });
                });

            // Save current edit if user clicked away (use the PREVIOUS editing cell)
            if save_current_edit {
                if let Some((edit_row, edit_col)) = previous_editing_cell {
                    if let Some(row_data) = self.data.get_mut(edit_row) {
                        if let Some(edit_cell) = row_data.get_mut(edit_col) {
                            *edit_cell = self.edit_buffer.clone();
                            self.has_unsaved_changes = true;
                        }
                    }
                }
            }

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
            if let Some((row_idx, col_idx)) = clear_cell {
                self.save_undo_state();
                if let Some(row_data) = self.data.get_mut(row_idx) {
                    if let Some(cell) = row_data.get_mut(col_idx) {
                        cell.clear();
                    }
                }
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
                    }); // End of ScrollArea
            }); // End of add_enabled_ui
        });

        // Draw unified confirmation modal
        if self.pending_action != PendingAction::None {
            // Check for Escape key to close modal
            if ctx.input(|i| i.key_pressed(egui::Key::Escape)) {
                self.pending_action = PendingAction::None;
            }

            let (title, message, confirm_label) = match &self.pending_action {
                PendingAction::NewFile => (
                    "Confirm New File",
                    "Are you sure you want to create a new file?",
                    "Yes, create new file"
                ),
                PendingAction::OpenFile => (
                    "Confirm Open File",
                    "Are you sure you want to open a file?",
                    "Yes, open file"
                ),
                PendingAction::Exit => (
                    "Confirm Exit",
                    "Are you sure you want to exit?",
                    "Yes, exit"
                ),
                PendingAction::None => ("", "", ""),
            };

            egui::Window::new(title)
                .collapsible(false)
                .resizable(false)
                .anchor(egui::Align2::CENTER_CENTER, [0.0, 0.0])
                .show(ctx, |ui| {
                    ui.label(message);
                    ui.label("All unsaved changes will be lost.");
                    ui.add_space(10.0);

                    ui.horizontal(|ui| {
                        if ui.button(confirm_label).clicked() {
                            match self.pending_action {
                                PendingAction::NewFile => {
                                    self.data = vec![vec![String::new(); 10]; 20];
                                    self.file_path = None;
                                    self.has_unsaved_changes = false;
                                    self.pending_action = PendingAction::None;
                                }
                                PendingAction::OpenFile => {
                                    #[cfg(not(target_arch = "wasm32"))]
                                    {
                                        if let Some(path) = rfd::FileDialog::new()
                                            .add_filter("CSV", &["csv"])
                                            .pick_file()
                                        {
                                            self.load_csv(path);
                                        }
                                    }
                                    #[cfg(target_arch = "wasm32")]
                                    {
                                        self.trigger_open_file();
                                    }
                                    self.pending_action = PendingAction::None;
                                }
                                PendingAction::Exit => {
                                    // Set allowed_to_close so the next close attempt succeeds
                                    self.allowed_to_close = true;
                                    self.pending_action = PendingAction::None;
                                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                                }
                                PendingAction::None => {}
                            }
                        }
                        if ui.button("Cancel").clicked() {
                            self.pending_action = PendingAction::None;
                            self.allowed_to_close = false;
                        }
                    });
                });
        }

        // Search window
        if self.search_window_open {
            // Handle Escape key to close search window
            if ctx.input(|i| i.key_pressed(egui::Key::Escape)) {
                self.search_window_open = false;
            }

            let mut window_open = true;
            egui::Window::new("Search")
                .open(&mut window_open)
                .collapsible(false)
                .resizable(false)
                .default_width(400.0)
                .show(ctx, |ui| {
                    ui.horizontal(|ui| {
                        ui.label("Find:");
                        let response = ui.text_edit_singleline(&mut self.search_query);

                        // Auto-focus the text field when window opens
                        if self.search_window_open {
                            response.request_focus();
                        }

                        // Enter key: search if no results yet, otherwise go to next result
                        if ui.input(|i| i.key_pressed(egui::Key::Enter)) && response.has_focus() {
                            if self.search_results.is_empty() {
                                self.perform_search();
                            } else {
                                self.go_to_next_search_result();
                            }
                        }
                    });

                    ui.horizontal(|ui| {
                        ui.checkbox(&mut self.search_case_sensitive, "Case sensitive");

                        if ui.button("Search").clicked() {
                            self.perform_search();
                        }
                    });

                    ui.separator();

                    if !self.search_results.is_empty() {
                        ui.label(format!(
                            "Found {} match{} (showing {} of {})",
                            self.search_results.len(),
                            if self.search_results.len() == 1 { "" } else { "es" },
                            self.current_search_result + 1,
                            self.search_results.len()
                        ));

                        ui.horizontal(|ui| {
                            if ui.button("Previous").clicked() {
                                self.go_to_prev_search_result();
                            }
                            if ui.button("Next").clicked() {
                                self.go_to_next_search_result();
                            }
                        });
                    } else if !self.search_query.is_empty() {
                        ui.label("No matches found");
                    }
                });

            if !window_open {
                self.search_window_open = false;
            }
        }
    }
}
