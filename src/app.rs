use std::path::PathBuf;

use anyhow::Result as AnyResult;
use eframe::{App, CreationContext, egui};
use egui::{Color32, FontData, FontDefinitions, FontFamily, RichText, TextEdit};
use poll_promise::Promise;
use rfd::FileDialog;

use crate::excel::{SplitResult, split_excel_file};

#[cfg(not(target_arch = "wasm32"))]
use std::fs;

pub struct ExcelHelperApp {
    header_row_input: String,
    row_count_input: String,
    selected_file: Option<PathBuf>,
    status: StatusMessage,
    split_promise: Option<Promise<AnyResult<SplitResult>>>,
    fonts_configured: bool,
}

impl Default for ExcelHelperApp {
    fn default() -> Self {
        Self {
            header_row_input: "1".into(),
            row_count_input: "500".into(),
            selected_file: None,
            status: StatusMessage::Idle,
            split_promise: None,
            fonts_configured: false,
        }
    }
}

impl ExcelHelperApp {
    pub fn new(cc: &CreationContext<'_>) -> Self {
        let mut app = Self::default();
        app.ensure_fonts(&cc.egui_ctx);
        app
    }

    fn ensure_fonts(&mut self, ctx: &egui::Context) {
        if self.fonts_configured {
            return;
        }

        if let Some(bytes) = load_cjk_font() {
            let mut fonts = FontDefinitions::default();
            fonts
                .font_data
                .insert("cjk_fallback".into(), FontData::from_owned(bytes));

            fonts
                .families
                .get_mut(&FontFamily::Proportional)
                .unwrap()
                .insert(0, "cjk_fallback".into());
            fonts
                .families
                .get_mut(&FontFamily::Monospace)
                .unwrap()
                .insert(0, "cjk_fallback".into());

            ctx.set_fonts(fonts);
        }

        self.fonts_configured = true;
    }

    fn pick_file(&mut self) {
        if let Some(path) = FileDialog::new()
            .add_filter("Excel 文件", &["xlsx", "xls"])
            .set_title("选择需要拆分的 Excel 文件")
            .pick_file()
        {
            self.selected_file = Some(path);
        }
    }

    fn start_split(&mut self) {
        if self.split_promise.is_some() {
            return;
        }

        let path = match self.selected_file.clone() {
            Some(path) => path,
            None => {
                self.status = StatusMessage::error("请先选择 Excel 文件");
                return;
            }
        };

        let header_rows = match self.parse_header_rows() {
            Ok(value) => value,
            Err(msg) => {
                self.status = StatusMessage::error(msg);
                return;
            }
        };

        let row_limit = match self.parse_row_limit() {
            Ok(value) => value,
            Err(msg) => {
                self.status = StatusMessage::error(msg);
                return;
            }
        };

        if row_limit <= header_rows {
            self.status = StatusMessage::error("拆分行数必须大于表头行数");
            return;
        }

        let promise = Promise::spawn_thread("excel-split", move || {
            split_excel_file(&path, row_limit, header_rows)
        });
        self.split_promise = Some(promise);
        self.status = StatusMessage::info("正在拆分，请稍候...");
    }

    fn parse_header_rows(&self) -> Result<usize, String> {
        let trimmed = self.header_row_input.trim();
        if trimmed.is_empty() {
            return Err("请输入表头行数".into());
        }

        let value: usize = trimmed
            .parse()
            .map_err(|_| "表头行数必须是正整数".to_string())?;
        if value == 0 {
            return Err("表头行数必须大于 0".into());
        }

        Ok(value)
    }

    fn parse_row_limit(&self) -> Result<usize, String> {
        let trimmed = self.row_count_input.trim();
        if trimmed.is_empty() {
            return Err("请输入拆分行数".into());
        }

        let value: usize = trimmed
            .parse()
            .map_err(|_| "行数必须是正整数".to_string())?;
        if value == 0 {
            return Err("行数必须大于 0".into());
        }

        Ok(value)
    }

    fn poll_promise(&mut self) {
        if let Some(promise) = self.split_promise.take() {
            match promise.try_take() {
                Ok(result) => match result {
                    Ok(split_result) => self.handle_success(split_result),
                    Err(err) => self.status = StatusMessage::error(format!("拆分失败: {err}")),
                },
                Err(promise) => {
                    self.split_promise = Some(promise);
                }
            }
        }
    }

    fn handle_success(&mut self, summary: SplitResult) {
        let mut message = format!(
            "拆分完成，共 {} 行（其中表头 {} 行）。\n生成 {} 个文件：",
            summary.total_rows,
            summary.header_rows,
            summary.chunks.len()
        );

        for (idx, chunk) in summary.chunks.iter().enumerate() {
            message.push_str(&format!(
                "\n第{}部分: {} 行（数据 {} 行） -> {}",
                idx + 1,
                chunk.total_rows,
                chunk.data_rows,
                chunk.file_path.display()
            ));
        }

        self.status = StatusMessage::success(message);
    }
}

impl App for ExcelHelperApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.ensure_fonts(ctx);
        self.poll_promise();

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Excel 拆分助手");
            ui.label("输入表头行数和单个文件的最大行数，程序会把表格按需拆成多个文件并保留表头。");
            ui.separator();

            ui.horizontal(|ui| {
                ui.label("表头行数：");
                let edit = TextEdit::singleline(&mut self.header_row_input)
                    .hint_text("例如 2")
                    .desired_width(120.0);
                ui.add(edit);
            });

            ui.horizontal(|ui| {
                ui.label("拆分行数：");
                let edit = TextEdit::singleline(&mut self.row_count_input)
                    .hint_text("例如 500")
                    .desired_width(120.0);
                ui.add(edit);
            });

            ui.horizontal_wrapped(|ui| {
                ui.label("目标文件：");
                let label_text = self
                    .selected_file
                    .as_ref()
                    .map(|p| p.display().to_string())
                    .unwrap_or_else(|| "尚未选择".into());
                ui.label(RichText::new(label_text).monospace());

                if ui.button("选择 Excel 文件").clicked() {
                    self.pick_file();
                }
            });

            let busy = self.split_promise.is_some();
            let button = ui.add_enabled(!busy, egui::Button::new("拆分 (Split)"));
            if button.clicked() {
                self.start_split();
            }

            if busy {
                ui.add_space(4.0);
                ui.horizontal(|ui| {
                    ui.spinner();
                    ui.label("正在处理大文件，请稍候...");
                });
            }

            if let Some((color, text)) = self.status.display() {
                ui.add_space(8.0);
                ui.colored_label(color, text);
            }
        });
    }
}

#[derive(Debug, Clone)]
enum StatusMessage {
    Idle,
    Info(String),
    Success(String),
    Error(String),
}

impl StatusMessage {
    fn info<S: Into<String>>(msg: S) -> Self {
        Self::Info(msg.into())
    }

    fn success<S: Into<String>>(msg: S) -> Self {
        Self::Success(msg.into())
    }

    fn error<S: Into<String>>(msg: S) -> Self {
        Self::Error(msg.into())
    }

    fn display(&self) -> Option<(Color32, &str)> {
        match self {
            StatusMessage::Idle => None,
            StatusMessage::Info(msg) => Some((Color32::LIGHT_BLUE, msg.as_str())),
            StatusMessage::Success(msg) => Some((Color32::from_rgb(0, 128, 0), msg.as_str())),
            StatusMessage::Error(msg) => Some((Color32::RED, msg.as_str())),
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn load_cjk_font() -> Option<Vec<u8>> {
    #[cfg(target_os = "windows")]
    {
        const CANDIDATES: [&str; 3] = [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/msyh.ttf",
            "C:/Windows/Fonts/simhei.ttf",
        ];

        for path in CANDIDATES {
            if let Ok(bytes) = fs::read(path) {
                return Some(bytes);
            }
        }
    }

    #[cfg(target_os = "macos")]
    {
        const CANDIDATES: [&str; 2] = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
        ];

        for path in CANDIDATES {
            if let Ok(bytes) = fs::read(path) {
                return Some(bytes);
            }
        }
    }

    #[cfg(target_os = "linux")]
    {
        const CANDIDATES: [&str; 2] = [
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ];

        for path in CANDIDATES {
            if let Ok(bytes) = fs::read(path) {
                return Some(bytes);
            }
        }
    }

    None
}

#[cfg(target_arch = "wasm32")]
fn load_cjk_font() -> Option<Vec<u8>> {
    None
}
