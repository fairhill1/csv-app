use eframe::egui;
use std::collections::HashMap;
use std::path::PathBuf;

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
}

impl Default for SpreadsheetApp {
    fn default() -> Self {
        Self {
            data: Vec::new(),
            file_path: None,
            editing_cell: None,
            edit_buffer: String::new(),
            column_widths: HashMap::new(),
            default_column_width: 120.0,
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

    fn get_column_width(&self, col_idx: usize) -> f32 {
        *self.column_widths.get(&col_idx).unwrap_or(&self.default_column_width)
    }

    fn set_column_width(&mut self, col_idx: usize, width: f32) {
        self.column_widths.insert(col_idx, width.max(30.0));
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
        let cols = if let Some(first_row) = self.data.first() {
            first_row.len()
        } else {
            10
        };
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
}

impl eframe::App for SpreadsheetApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
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
            egui::ScrollArea::both().show(ui, |ui| {
                if self.data.is_empty() {
                    ui.label("No data loaded. Create a new spreadsheet or open a CSV file.");
                    return;
                }

                let num_cols = self.data.iter().map(|row| row.len()).max().unwrap_or(0);
                let row_height = 25.0;
                let row_label_width = 50.0;

                // Pre-compute column widths to avoid borrow conflicts
                let col_widths: Vec<f32> = (0..num_cols)
                    .map(|i| self.get_column_width(i))
                    .collect();

                ui.style_mut().spacing.item_spacing = egui::vec2(0.0, 0.0);

                // Header row with column letters and resize handles
                ui.horizontal(|ui| {
                    // Empty corner cell
                    ui.add_sized([row_label_width, row_height], egui::Label::new(""));

                    for col_idx in 0..num_cols {
                        let col_width = col_widths[col_idx];

                        // Column header with resize handle on the right edge
                        let (rect, _response) = ui.allocate_exact_size(
                            egui::vec2(col_width, row_height),
                            egui::Sense::hover()
                        );

                        ui.painter().text(
                            rect.center(),
                            egui::Align2::CENTER_CENTER,
                            Self::col_index_to_letter(col_idx),
                            egui::FontId::default(),
                            ui.visuals().text_color()
                        );

                        // Resize handle at right edge
                        let handle_rect = egui::Rect::from_min_size(
                            egui::pos2(rect.right() - 4.0, rect.top()),
                            egui::vec2(8.0, row_height)
                        );
                        let handle_response = ui.interact(
                            handle_rect,
                            ui.id().with(("resize", col_idx)),
                            egui::Sense::drag()
                        );

                        if handle_response.dragged() {
                            let new_width = col_width + handle_response.drag_delta().x;
                            self.set_column_width(col_idx, new_width);
                        }

                        if handle_response.hovered() || handle_response.dragged() {
                            ui.ctx().set_cursor_icon(egui::CursorIcon::ResizeColumn);
                            ui.painter().vline(
                                rect.right(),
                                (rect.top())..=(rect.bottom()),
                                egui::Stroke::new(2.0, egui::Color32::from_gray(150))
                            );
                        } else {
                            ui.painter().vline(
                                rect.right(),
                                (rect.top())..=(rect.bottom()),
                                egui::Stroke::new(1.0, ui.visuals().widgets.noninteractive.bg_stroke.color)
                            );
                        }
                    }
                });

                ui.separator();

                // Data rows
                for (row_idx, row) in self.data.iter_mut().enumerate() {
                    let bg_color = if row_idx % 2 == 0 {
                        ui.visuals().faint_bg_color
                    } else {
                        egui::Color32::TRANSPARENT
                    };

                    ui.horizontal(|ui| {
                        // Row number label
                        let (rect, _) = ui.allocate_exact_size(
                            egui::vec2(row_label_width, row_height),
                            egui::Sense::hover()
                        );
                        if bg_color != egui::Color32::TRANSPARENT {
                            ui.painter().rect_filled(rect, 0.0, bg_color);
                        }
                        ui.painter().text(
                            rect.center(),
                            egui::Align2::CENTER_CENTER,
                            format!("{}", row_idx + 1),
                            egui::FontId::default(),
                            ui.visuals().text_color()
                        );

                        // Ensure row has enough columns
                        while row.len() < num_cols {
                            row.push(String::new());
                        }

                        for (col_idx, cell) in row.iter_mut().enumerate() {
                            let col_width = col_widths[col_idx];
                            let cell_id = (row_idx, col_idx);
                            let is_editing = self.editing_cell == Some(cell_id);

                            let (cell_rect, response) = ui.allocate_exact_size(
                                egui::vec2(col_width, row_height),
                                egui::Sense::click()
                            );

                            if bg_color != egui::Color32::TRANSPARENT {
                                ui.painter().rect_filled(cell_rect, 0.0, bg_color);
                            }

                            // Draw cell border
                            ui.painter().rect_stroke(
                                cell_rect,
                                0.0,
                                egui::Stroke::new(1.0, ui.visuals().widgets.noninteractive.bg_stroke.color),
                                egui::epaint::StrokeKind::Outside
                            );

                            if is_editing {
                                let text_edit = egui::TextEdit::singleline(&mut self.edit_buffer)
                                    .frame(false);

                                let edit_response = ui.put(cell_rect.shrink(2.0), text_edit);

                                if edit_response.lost_focus() {
                                    *cell = self.edit_buffer.clone();
                                    self.editing_cell = None;
                                }

                                edit_response.request_focus();
                            } else {
                                // Clip text to cell boundaries
                                let text_rect = cell_rect.shrink2(egui::vec2(4.0, 0.0));
                                ui.painter().with_clip_rect(text_rect).text(
                                    cell_rect.left_center() + egui::vec2(4.0, 0.0),
                                    egui::Align2::LEFT_CENTER,
                                    cell.as_str(),
                                    egui::FontId::default(),
                                    ui.visuals().text_color()
                                );

                                if response.clicked() {
                                    self.editing_cell = Some(cell_id);
                                    self.edit_buffer = cell.clone();
                                }

                                response.context_menu(|ui| {
                                    if ui.button("Clear").clicked() {
                                        cell.clear();
                                        ui.close();
                                    }
                                });
                            }
                        }
                    });
                }
            });
        });
    }
}
