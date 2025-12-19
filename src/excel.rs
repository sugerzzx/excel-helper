use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, anyhow};
use calamine::{Data, Reader, open_workbook_auto};
use quick_xml::{Reader as XmlReader, events::Event};
use rust_xlsxwriter::{Format, Workbook};
use zip::ZipArchive;

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

#[derive(Debug, Clone)]
struct MergeRange {
    start_row: usize,
    end_row: usize,
    start_col: usize,
    end_col: usize,
}

struct ChunkMerge {
    start_row: u32,
    end_row: u32,
    start_col: u16,
    end_col: u16,
    value: String,
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

    let merge_ranges = extract_merge_ranges(source, &sheet_name)?;

    let data_capacity = chunk_size - header_rows;
    let mut chunks = Vec::new();

    if data_rows.is_empty() {
        let path = build_output_path(source, 1);
        let chunk_merges = map_chunk_merges(&merge_ranges, header_rows, 0, 0, &header, &data_rows);
        write_chunk(&path, &header, &[], &chunk_merges)?;
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
            let chunk_merges =
                map_chunk_merges(&merge_ranges, header_rows, start, end, &header, &data_rows);
            write_chunk(&path, &header, chunk_data, &chunk_merges)?;
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
    merges: &[ChunkMerge],
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

    if !merges.is_empty() {
        let merge_format = Format::new();
        for merge in merges {
            worksheet.merge_range(
                merge.start_row,
                merge.start_col,
                merge.end_row,
                merge.end_col,
                &merge.value,
                &merge_format,
            )?;
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

fn extract_merge_ranges(source: &Path, sheet_name: &str) -> Result<Vec<MergeRange>> {
    let extension = source
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    if extension != "xlsx" {
        return Ok(Vec::new());
    }

    let file = File::open(source)
        .with_context(|| format!("无法以 ZIP 方式打开 Excel 文件: {}", source.display()))?;
    let mut archive = ZipArchive::new(file)
        .with_context(|| "无法解压 Excel 文件以读取合并单元信息".to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let rel_id = find_sheet_rel_id(&workbook_xml, sheet_name)?;
    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let target = find_sheet_target(&rels_xml, &rel_id)?;
    let full_path = format!("xl/{}", target.trim_start_matches('/'));
    let sheet_xml = read_zip_entry(&mut archive, &full_path)?;
    Ok(parse_merge_cells(&sheet_xml))
}

fn read_zip_entry<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> Result<String> {
    let mut file = archive
        .by_name(path)
        .with_context(|| format!("Excel 文件缺少 {path}"))?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)
        .with_context(|| format!("无法读取 {path}"))?;
    Ok(contents)
}

fn find_sheet_rel_id(workbook_xml: &str, sheet_name: &str) -> Result<String> {
    let mut reader = XmlReader::from_str(workbook_xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    while let Ok(event) = reader.read_event_into(&mut buf) {
        match event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                if !e.name().as_ref().ends_with(b"sheet") {
                    continue;
                }
                let mut name_attr = None;
                let mut rel_id = None;
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    if key.ends_with(b"name") {
                        name_attr = Some(attr.decode_and_unescape_value(&reader)?.into_owned());
                    } else if key.ends_with(b":id") {
                        rel_id = Some(attr.decode_and_unescape_value(&reader)?.into_owned());
                    }
                }
                if let (Some(name), Some(rid)) = (name_attr, rel_id) {
                    if name == sheet_name {
                        return Ok(rid);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Err(anyhow!(
        "无法在 workbook.xml 中找到工作表 {sheet_name} 的关系信息"
    ))
}

fn find_sheet_target(rels_xml: &str, rel_id: &str) -> Result<String> {
    let mut reader = XmlReader::from_str(rels_xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    while let Ok(event) = reader.read_event_into(&mut buf) {
        match event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                if !e.name().as_ref().ends_with(b"Relationship") {
                    continue;
                }
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    if key.ends_with(b"Id") {
                        id = Some(attr.decode_and_unescape_value(&reader)?.into_owned());
                    } else if key.ends_with(b"Target") {
                        target = Some(attr.decode_and_unescape_value(&reader)?.into_owned());
                    }
                }
                if id.as_deref() == Some(rel_id) {
                    if let Some(target) = target {
                        return Ok(target);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Err(anyhow!("无法定位工作表的 XML 路径"))
}

fn parse_merge_cells(sheet_xml: &str) -> Vec<MergeRange> {
    let mut reader = XmlReader::from_str(sheet_xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut ranges = Vec::new();
    while let Ok(event) = reader.read_event_into(&mut buf) {
        match event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                if !e.name().as_ref().ends_with(b"mergeCell") {
                    continue;
                }
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref().ends_with(b"ref") {
                        if let Ok(value) = attr.decode_and_unescape_value(&reader) {
                            if let Some(range) = parse_range_ref(value.trim()) {
                                ranges.push(range);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    ranges
}

fn parse_range_ref(range: &str) -> Option<MergeRange> {
    let mut parts = range.split(':');
    let start = parts.next()?.trim();
    let end = parts.next().unwrap_or(start).trim();
    let (start_row, start_col) = parse_cell_ref(start)?;
    let (end_row, end_col) = parse_cell_ref(end)?;
    let (row_min, row_max) = if start_row <= end_row {
        (start_row, end_row)
    } else {
        (end_row, start_row)
    };
    let (col_min, col_max) = if start_col <= end_col {
        (start_col, end_col)
    } else {
        (end_col, start_col)
    };
    Some(MergeRange {
        start_row: row_min,
        end_row: row_max,
        start_col: col_min,
        end_col: col_max,
    })
}

fn parse_cell_ref(cell: &str) -> Option<(usize, usize)> {
    if cell.is_empty() {
        return None;
    }
    let mut col_part = String::new();
    let mut row_part = String::new();
    for ch in cell.chars() {
        if ch.is_ascii_alphabetic() {
            if !row_part.is_empty() {
                return None;
            }
            col_part.push(ch);
        } else if ch.is_ascii_digit() {
            row_part.push(ch);
        }
    }
    if col_part.is_empty() || row_part.is_empty() {
        return None;
    }
    let col = column_label_to_index(&col_part)?;
    let row = row_part.parse::<usize>().ok()?.checked_sub(1)?;
    Some((row, col))
}

fn column_label_to_index(label: &str) -> Option<usize> {
    let mut value = 0usize;
    for ch in label.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let offset = (ch.to_ascii_uppercase() as u8).checked_sub(b'A')? as usize;
        value = value.checked_mul(26)? + offset + 1;
    }
    value.checked_sub(1)
}

fn map_chunk_merges(
    merges: &[MergeRange],
    header_rows: usize,
    data_start: usize,
    data_end: usize,
    header_data: &[Vec<String>],
    data_data: &[Vec<String>],
) -> Vec<ChunkMerge> {
    let mut result = Vec::new();
    for merge in merges {
        if !row_in_chunk(merge.start_row, header_rows, data_start, data_end)
            || !row_in_chunk(merge.end_row, header_rows, data_start, data_end)
        {
            continue;
        }
        let start_row = map_row_to_chunk(merge.start_row, header_rows, data_start);
        let end_row = map_row_to_chunk(merge.end_row, header_rows, data_start);
        let start_col = match u16::try_from(merge.start_col) {
            Ok(col) => col,
            Err(_) => continue,
        };
        let end_col = match u16::try_from(merge.end_col) {
            Ok(col) => col,
            Err(_) => continue,
        };
        let value = get_cell_value(
            header_data,
            data_data,
            header_rows,
            merge.start_row,
            merge.start_col,
        );
        result.push(ChunkMerge {
            start_row: start_row as u32,
            end_row: end_row as u32,
            start_col,
            end_col,
            value,
        });
    }
    result
}

fn row_in_chunk(row: usize, header_rows: usize, data_start: usize, data_end: usize) -> bool {
    if row < header_rows {
        true
    } else {
        let data_idx = row - header_rows;
        data_idx >= data_start && data_idx < data_end
    }
}

fn map_row_to_chunk(row: usize, header_rows: usize, data_start: usize) -> usize {
    if row < header_rows {
        row
    } else {
        header_rows + (row - header_rows - data_start)
    }
}

fn get_cell_value(
    header_rows: &[Vec<String>],
    data_rows: &[Vec<String>],
    header_len: usize,
    row: usize,
    col: usize,
) -> String {
    if row < header_len {
        header_rows
            .get(row)
            .and_then(|r| r.get(col))
            .cloned()
            .unwrap_or_default()
    } else {
        let data_idx = row - header_len;
        data_rows
            .get(data_idx)
            .and_then(|r| r.get(col))
            .cloned()
            .unwrap_or_default()
    }
}
