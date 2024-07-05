use std::sync::{Arc, Mutex};
use rustler::{Atom, Env, NifResult, ResourceArc, Error as RustlerError, Term};

use umya_spreadsheet::reader;
use umya_spreadsheet::writer;
use umya_spreadsheet::reader::xlsx::XlsxError;
use umya_spreadsheet::structs::Spreadsheet;
use umya_spreadsheet::helper::coordinate::index_from_coordinate;

#[allow(dead_code)]
struct SpreadsheetResource(Arc<Mutex<Spreadsheet>>);

struct WorksheetResource {
    spreadsheet: Arc<Mutex<Spreadsheet>>,
    sheet_name: String,
}

struct CellResource {
    spreadsheet: Arc<Mutex<Spreadsheet>>,
    sheet_name: String,
    cell_ref: String,
}

pub fn load(env: Env, _: Term) -> bool {
    rustler::resource!(SpreadsheetResource, env);
    rustler::resource!(WorksheetResource, env);
    rustler::resource!(CellResource, env);
    true
}

mod atoms {
    rustler::atoms! {
        ok,
        error,
        unknown // Other error
    }
}

fn xlsx_error_to_str(err: XlsxError) -> String {
    match err {
        XlsxError::Io(_) => "io_error".to_string(),
        XlsxError::Xml(_) => "xml_error".to_string(),
        XlsxError::Zip(_) => "zip_error".to_string(),
        XlsxError::Uft8(_) => "uft8_error".to_string(),
    }
}


#[rustler::nif]
pub fn read(path: String) -> NifResult<ResourceArc<SpreadsheetResource>> {
    let p = std::path::Path::new(&path);
    match reader::xlsx::read(p) {
        Ok(spreadsheet) => Ok(ResourceArc::new(SpreadsheetResource(Arc::new(Mutex::new(spreadsheet))))),
        Err(error) => return Err(RustlerError::Term(Box::new(xlsx_error_to_str(error)))),
    }
}


#[rustler::nif]
pub fn get_sheet(spreadsheet: ResourceArc<SpreadsheetResource>, index: usize) -> NifResult<Option<ResourceArc<WorksheetResource>>> {
    let spreadsheet_arc = spreadsheet.0.clone();
    let worksheet = {
        let spreadsheet = spreadsheet_arc.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
        spreadsheet.get_sheet(&index).map(|sheet| sheet.clone())
    };
    if let Some(sheet) = worksheet  {
        Ok(Some(ResourceArc::new(WorksheetResource {
            spreadsheet: spreadsheet_arc,
            sheet_name: sheet.get_name().to_string(),
        })))
    } else {
        Ok(None)
    }
}

#[rustler::nif]
pub fn get_sheet_by_name(spreadsheet: ResourceArc<SpreadsheetResource>, name: String) -> NifResult<Option<ResourceArc<WorksheetResource>>> {
    let spreadsheet_arc = spreadsheet.0.clone();
    let exists = {
        let spreadsheet = spreadsheet_arc.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
        spreadsheet.get_sheet_by_name(&name).is_some()
    };
    if exists {
        Ok(Some(ResourceArc::new(WorksheetResource {
            spreadsheet: spreadsheet_arc,
            sheet_name: name,
        })))
    } else {
        Ok(None)
    }
}

#[rustler::nif]
fn get_cell(worksheet: ResourceArc<WorksheetResource>, cell_ref: String) -> NifResult<Option<ResourceArc<CellResource>>> {
    let spreadsheet = worksheet.spreadsheet.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
    let sheet = spreadsheet.get_sheet_by_name(&worksheet.sheet_name).ok_or_else(|| rustler::Error::Term(Box::new("sheet_not_found")))?;
    if sheet.get_cell(cell_ref.as_str()).is_some() {
        Ok(Some(ResourceArc::new(CellResource {
            spreadsheet: worksheet.spreadsheet.clone(),
            sheet_name: worksheet.sheet_name.clone(),
            cell_ref,
        })))
    } else {
        Ok(None)
    }
}

#[rustler::nif(name = "get_cell_value")]
fn get_cell_value_by_cell(cell: ResourceArc<CellResource>) -> NifResult<Option<String>> {
    let spreadsheet = cell.spreadsheet.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
    let sheet = spreadsheet.get_sheet_by_name(&cell.sheet_name).ok_or_else(|| rustler::Error::Term(Box::new("sheet_not_found")))?;
    let cell = sheet.get_cell(cell.cell_ref.as_str()).ok_or_else(|| rustler::Error::Term(Box::new("cell_not_found")))?;
    let value = cell.get_value().to_string();
    if value.is_empty() {
        Ok(None)
    } else {
        Ok(Some(value))
    }
}

#[rustler::nif(name = "get_cell_value")]
fn get_cell_value_by_sheet(worksheet: ResourceArc<WorksheetResource>, cell_ref: String) -> NifResult<Option<String>> {
    let spreadsheet = worksheet.spreadsheet.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
    let sheet = spreadsheet.get_sheet_by_name(&worksheet.sheet_name).ok_or_else(|| rustler::Error::Term(Box::new("sheet_not_found")))?;

    let (col, row, _, _) = index_from_coordinate(cell_ref.as_str());
    if col.is_none() || row.is_none() {
        return Ok(None);
    } else {
        let cell = sheet.get_cell(cell_ref.as_str())
            .ok_or_else(|| rustler::Error::Term(Box::new("cell_not_found")))?;
        let value = cell.get_value().to_string();
        if value.is_empty() {
            Ok(None)
        } else {
            Ok(Some(value))
        }
    }
}

#[rustler::nif(name = "set_cell_value")]
fn set_cell_value_by_cell(cell: ResourceArc<CellResource>, value: String) -> NifResult<Atom> {
    let mut spreadsheet = cell.spreadsheet.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
    let sheet = spreadsheet.get_sheet_by_name_mut(&cell.sheet_name).ok_or_else(|| rustler::Error::Term(Box::new("sheet_not_found")))?;
    sheet.get_cell_mut(cell.cell_ref.as_str()).set_value(value);
    Ok(atoms::ok())
}

#[rustler::nif(name = "set_cell_value")]
fn set_cell_value_by_sheet(worksheet: ResourceArc<WorksheetResource>, cell_ref: String, value: String) -> NifResult<Atom> {
    let spreadsheet = worksheet.spreadsheet.clone();
    let mut spreadsheet_lock = spreadsheet.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
    let sheet = spreadsheet_lock.get_sheet_by_name_mut(&worksheet.sheet_name).ok_or_else(|| rustler::Error::Term(Box::new("sheet_not_found")))?;
    
    // Always write value to the cell even if that cell is not existed.
    sheet.get_cell_mut(cell_ref.as_str()).set_value(&value);
    Ok(atoms::ok())
}

#[rustler::nif]
pub fn save(spreadsheet: ResourceArc<SpreadsheetResource>, path: String) -> NifResult<Atom> {
    let spreadsheet = spreadsheet.0.lock().map_err(|_| rustler::Error::Term(Box::new("lock_failed")))?;
    let p = std::path::Path::new(&path);
    writer::xlsx::write(&spreadsheet, p).map_err(|e| rustler::Error::Term(Box::new(format!("save_failed: {}", e))))?;
    Ok(atoms::ok())
}

rustler::init!(
    "Elixir.ExcelixirRustler",
    [
        read,
        get_sheet,
        get_sheet_by_name,
        get_cell,
        get_cell_value_by_cell,
        get_cell_value_by_sheet,
        set_cell_value_by_cell,
        set_cell_value_by_sheet,
        save
    ],
    load = load
);
