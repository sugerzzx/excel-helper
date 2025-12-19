use std::path::{Path, PathBuf};

use anyhow::{Context, Result, anyhow};
use calamine::{Data, Reader, open_workbook_auto};
use rust_xlsxwriter::Workbook;

/// Metadata describing the generated files and helpful stats for the UI.
pub struct SplitResult {
    pub total_rows: usize,
    pub header_rows: usize,
    pub chunks: Vec<SplitChunk>,
}

/// Metadata for a single output file.
pub struct SplitChunk {
    pub file_path: PathBuf,
    pub total_rows: usize,
    pub data_rows: usize,
}

/// Splits the first worksheet of the given Excel file into multiple files while keeping the header.
pub fn split_excel_file(
    source: &Path,
    chunk_size: usize,
    header_rows: usize,
) -> Result<SplitResult> {
    if chunk_size == 0 {
        return Err(anyhow!("拆分的行数必须大于 0"));
    }

    if header_rows == 0 {
        return Err(anyhow!("表头行数必须大于 0"));
    }

    if chunk_size <= header_rows {
        return Err(anyhow!("拆分行数必须大于表头行数"));
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

    let rows: Vec<Vec<String>> = range.rows().map(convert_row).collect();
    if rows.len() < header_rows {
        return Err(anyhow!("工作表的行数小于指定的表头行数"));
    }

    let header = rows[..header_rows].to_vec();
    let data_rows = rows[header_rows..].to_vec();
    let total_rows = rows.len();

    let data_capacity = chunk_size - header_rows;
    let mut chunks = Vec::new();

    if data_rows.is_empty() {
        let path = build_output_path(source, 1);
        write_chunk(&path, &header, &[])?;
        chunks.push(SplitChunk {
            file_path: path,
            total_rows: header_rows,
            data_rows: 0,
        });
    } else {
        let mut start = 0;
        let mut index = 1;
        while start < data_rows.len() {
            let end = (start + data_capacity).min(data_rows.len());
            let chunk_data = &data_rows[start..end];
            let path = build_output_path(source, index);
            write_chunk(&path, &header, chunk_data)?;
            chunks.push(SplitChunk {
                file_path: path,
                total_rows: header_rows + chunk_data.len(),
                data_rows: chunk_data.len(),
            });
            start = end;
            index += 1;
        }
    }

    Ok(SplitResult {
        total_rows,
        header_rows,
        chunks,
    })
}

fn write_chunk(
    destination: &Path,
    header_rows: &[Vec<String>],
    data_rows: &[Vec<String>],
) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let mut current_row: u32 = 0;

    for header_row in header_rows {
        for (col_idx, value) in header_row.iter().enumerate() {
            worksheet.write_string(current_row, col_idx as u16, value)?;
        }
        current_row += 1;
    }

    for data_row in data_rows {
        for (col_idx, value) in data_row.iter().enumerate() {
            worksheet.write_string(current_row, col_idx as u16, value)?;
        }
        current_row += 1;
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

fn build_output_path(source: &Path, index: usize) -> PathBuf {
    let parent = source
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let stem = source
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("split");
    parent.join(format!("{stem}_part{index}.xlsx"))
}
