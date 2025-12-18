#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod app;
mod excel;

use app::ExcelHelperApp;
use eframe::{NativeOptions, egui};

fn main() -> eframe::Result<()> {
    let native_options = NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([520.0, 360.0])
            .with_min_inner_size([420.0, 300.0]),
        ..Default::default()
    };

    eframe::run_native(
        "Excel Helper",
        native_options,
        Box::new(|cc| Box::new(ExcelHelperApp::new(cc))),
    )
}
