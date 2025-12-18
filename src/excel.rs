use std::path::{Path, PathBuf};

use anyhow::{Context, Result, anyhow};
use calamine::{Data, Reader, open_workbook_auto};
use rust_xlsxwriter::Workbook;

/// Metadata that describes the generated files along with a few helpful stats.
pub struct SplitResult {
    pub first_file: PathBuf,
    pub second_file: PathBuf,
    pub total_rows: usize,
    pub first_segment_rows: usize,
    pub second_segment_rows: usize,
}

/// Splits the first worksheet of the given Excel file into two new files while keeping the header.
pub fn split_excel_file(source: &Path, chunk_size: usize) -> Result<SplitResult> {
    if chunk_size == 0 {
        return Err(anyhow!("拆分的行数必须大于 0"));
    }

    let mut workbook = open_workbook_auto(source)
        .with_context(|| format!("无法打开 Excel 文件: {}", source.display()))?;

    let sheet_name = workbook
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| anyhow!("所选文件中没有任何工作表"))?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .with_context(|| format!("无法读取工作表 {sheet_name}"))?;

    let mut rows = range.rows();
    let header_row = rows.next().ok_or_else(|| anyhow!("工作表为空，缺少表头"))?;
    let header = convert_row(header_row);

    let data_rows: Vec<Vec<String>> = rows.map(convert_row).collect();
    let total_rows = data_rows.len();

    let split_index = chunk_size.min(total_rows);
    let (first_segment, second_segment) = data_rows.split_at(split_index);

    let first_path = build_output_path(source, "part1");
    let second_path = build_output_path(source, "part2");

    write_segment(&first_path, &header, first_segment)?;
    write_segment(&second_path, &header, second_segment)?;

    Ok(SplitResult {
        first_file: first_path,
        second_file: second_path,
        total_rows,
        first_segment_rows: first_segment.len(),
        second_segment_rows: second_segment.len(),
    })
}

fn write_segment(destination: &Path, header: &[String], rows: &[Vec<String>]) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for (col_idx, value) in header.iter().enumerate() {
        worksheet.write_string(0, col_idx as u16, value)?;
    }

    for (row_idx, row) in rows.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            worksheet.write_string((row_idx + 1) as u32, col_idx as u16, value)?;
        }
    }

    workbook.save(destination)?;
    Ok(())
}

fn convert_row(row: &[Data]) -> Vec<String> {
    row.iter().map(format_cell).collect()
}

fn format_cell(value: &Data) -> String {
    match value {
        Data::Empty => String::new(),
        Data::String(s) => s.clone(),
        Data::Float(f) => format_float(*f),
        Data::Int(i) => i.to_string(),
        Data::Bool(b) => {
            if *b {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Data::DateTime(dt) => dt
            .as_datetime()
            .map(|date_time| date_time.format("%Y-%m-%d %H:%M:%S").to_string())
            .unwrap_or_else(|| dt.as_f64().to_string()),
        Data::DateTimeIso(iso) | Data::DurationIso(iso) => iso.clone(),
        Data::Error(e) => format!("错误: {e:?}"),
    }
}

fn format_float(value: f64) -> String {
    if !value.is_finite() {
        return value.to_string();
    }

    if value.fract() == 0.0 {
        format!("{:.0}", value)
    } else {
        let mut repr = format!("{value}");
        if let Some(point_pos) = repr.find('.') {
            while repr.ends_with('0') {
                repr.pop();
            }
            if repr.ends_with('.') {
                repr.push('0');
            }
            if repr.len() == point_pos {
                repr.push_str("0");
            }
        }
        repr
    }
}

fn build_output_path(source: &Path, suffix: &str) -> PathBuf {
    let parent = source
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let stem = source
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("split");
    parent.join(format!("{stem}_{suffix}.xlsx"))
}
