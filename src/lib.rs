use wasm_minimal_protocol::*;
use serde_json::{json, Value};
use umya_spreadsheet::*;
use std::io::Cursor;
use serde_json::Map;

initiate_protocol!();

#[wasm_func]
fn excel_to_json_values(file_bytes: &[u8], worksheet_name: &[u8]) -> Result<Vec<u8>, String> {
    let mutable_file_bytes = Cursor::new(file_bytes);
    let worksheet_name_string = std::str::from_utf8(worksheet_name).expect("Invalid UTF-8");

    let mut xl = umya_spreadsheet::reader::xlsx::read_reader(mutable_file_bytes,false).map_err(|e| e.to_string())?;

    xl.read_sheet_collection();

    let sheet = xl.get_sheet_by_name(worksheet_name_string).ok_or("Worksheet not found")?;

    let json_data = convert_sheet_to_json_values(sheet,&xl)?;

    Ok(json_data.to_string().into_bytes())
}

fn convert_sheet_to_json_values(sheet: &Worksheet,book:&Spreadsheet) -> Result<Value, String> {
    let mut cells_data = Vec::new();

    let merges = sheet.get_merge_cells();
    for cell in sheet.get_cell_collection() {
        let cell_coordinate = cell.get_coordinate();
        let cell_row = cell_coordinate.get_row_num();
        let cell_col = cell_coordinate.get_col_num();
    

        // Keep the cell if it's not within any merge range or if it's at the top-left of a merge range

            let mut tmp_cell_data = json!({
                "x": cell_col,
                "y": cell_row,
                "value": cell.get_value(),
            });
    
            cells_data.push(tmp_cell_data); 
        
    }

    // Sort the cells_data by row (y) and column (x) to maintain a consistent order
    cells_data.sort_by(|a, b| {
        let a_y = a["y"].as_i64().unwrap();
        let a_x = a["x"].as_i64().unwrap();
        let b_y = b["y"].as_i64().unwrap();
        let b_x = b["x"].as_i64().unwrap();
        (a_y, a_x).cmp(&(b_y, b_x))
    });

    Ok(json!({
        "cells": cells_data
    }))
}


#[wasm_func]
fn excel_to_json(file_bytes: &[u8], worksheet_name: &[u8]) -> Result<Vec<u8>, String> {
    let mutable_file_bytes = Cursor::new(file_bytes);
    let worksheet_name_string = std::str::from_utf8(worksheet_name).expect("Invalid UTF-8");

    let mut xl = umya_spreadsheet::reader::xlsx::read_reader(mutable_file_bytes,false).map_err(|e| e.to_string())?;

    xl.read_sheet_collection();

    let sheet = xl.get_sheet_by_name(worksheet_name_string).ok_or("Worksheet not found")?;

    let json_data = convert_sheet_to_json(sheet,&xl)?;

    Ok(json_data.to_string().into_bytes())
}

fn convert_sheet_to_json(sheet: &Worksheet,book:&Spreadsheet) -> Result<Value, String> {
    let mut cells_data = Vec::new();

    let merges = sheet.get_merge_cells();
    for cell in sheet.get_cell_collection() {
        let cell_coordinate = cell.get_coordinate();
        let cell_row = cell_coordinate.get_row_num();
        let cell_col = cell_coordinate.get_col_num();
    
        let mut found_range: Option<&Range> = None;

        // Check if the cell is within any merge range
        for range in merges.iter() {
            if cell_row >= range.get_coordinate_start_row().unwrap().get_num()
            && cell_row <= range.get_coordinate_end_row(  ).unwrap().get_num()
            && cell_col >= range.get_coordinate_start_col().unwrap().get_num()
            && cell_col <= range.get_coordinate_end_col(  ).unwrap().get_num()
            {
                found_range = Some(range);
                break; // Exit loop after finding the first matching range
            }
        }

        // Check if the cell is at the top-left of any merge range
        let is_top_left_of_merge = merges.iter().any(|range| {
               cell_row == range.get_coordinate_start_row().unwrap().get_num()
            && cell_col == range.get_coordinate_start_col().unwrap().get_num()
        });
        
        // Keep the cell if it's not within any merge range or if it's at the top-left of a merge range
        if found_range.is_none() || is_top_left_of_merge {
            let mut tmp_cell_data = json!({
                "x": cell_col,
                "y": cell_row,
                "value": cell.get_value(),
            });

            if let Some(varA) = cell.get_style().get_borders() {
                tmp_cell_data["boarders"] = json!({
                    "left": varA.get_left().get_style().get_value_string(),
                    "right": varA.get_right().get_style().get_value_string(),
                    "bottom": varA.get_bottom().get_style().get_value_string(),
                    "top": varA.get_top().get_style().get_value_string(),
                    "diag": varA.get_diagonal().get_style().get_value_string(),

                })
             }

             if let Some(varA) = cell.get_style().get_alignment() {
                tmp_cell_data["alignment"] = json!({
                    "horizontal": varA.get_horizontal().get_value_string(),
                    "vertical": varA.get_vertical().get_value_string(),
                    "rot": varA.get_text_rotation(),
                    "wrap": varA.get_wrap_text(),
                })
             }
            

            if is_top_left_of_merge{
                tmp_cell_data["w"] = json!(found_range.unwrap().get_coordinate_end_col(  ).unwrap().get_num() - 
                                           found_range.unwrap().get_coordinate_start_col().unwrap().get_num() + 1 );
                tmp_cell_data["h"] = json!(found_range.unwrap().get_coordinate_end_row(  ).unwrap().get_num() - 
                                           found_range.unwrap().get_coordinate_start_row().unwrap().get_num() + 1 );
            }
    
            if let Some(background_color) = cell.get_style().get_background_color() {
                if background_color.get_indexed() != &0  {
                    tmp_cell_data["fill_color"] = json!(INDEXED_COLORS.get(*background_color.get_indexed() as usize));
                } else if background_color.get_theme_index() != &0 {
                    tmp_cell_data["fill_color"] = json!(background_color.get_argb_with_theme(book.get_theme()));
                } else {
                    tmp_cell_data["fill_color"] = json!(background_color.get_argb().to_string());
                }
            }
    
            cells_data.push(tmp_cell_data); 
        }
    }

    let mut columns_data = Vec::new();
    for col in sheet.get_column_dimensions(){
        let tmp_col_data = json!({
            "x": col.get_col_num(),
            "width": col.get_width(),
            "hidden": col.get_hidden()
        });
        columns_data.push(tmp_col_data);
    }

    let mut rows_data = Vec::new(); 
    for row in sheet.get_row_dimensions(){
        let tmp_row_data = json!({
            "y": row.get_row_num(),
            "height": row.get_height(),
            "hidden": row.get_hidden()
        });
        rows_data.push(tmp_row_data);
    }


    // Sort the cells_data by row (y) and column (x) to maintain a consistent order
    cells_data.sort_by(|a, b| {
        let a_y = a["y"].as_i64().unwrap();
        let a_x = a["x"].as_i64().unwrap();
        let b_y = b["y"].as_i64().unwrap();
        let b_x = b["x"].as_i64().unwrap();
        (a_y, a_x).cmp(&(b_y, b_x))
    });


    Ok(json!({
        "columns":json!(columns_data),
        "rows":json!(rows_data),
        "cells":json!(cells_data)
    }))
}



#[wasm_func]
pub fn concatenate(arg1: &[u8], arg2: &[u8]) -> Vec<u8> {
    [arg1, b"*", arg2].concat()
}



// // I hate that this exist but i put it here. Its a workaround for the build process.
use getrandom::register_custom_getrandom;
use getrandom::Error;

// Some application-specific error code
pub fn always_fail(_buf: &mut [u8]) -> Result<(), Error> {
    Ok(())
}

register_custom_getrandom!(always_fail);


//Literally copied from umya since i couldn't figure out how to reference it.
const INDEXED_COLORS: &[&str] = &[
    "FF000000", //  System Colour #1 - Black
    "FFFFFFFF", //  System Colour #2 - White
    "FFFF0000", //  System Colour #3 - Red
    "FF00FF00", //  System Colour #4 - Green
    "FF0000FF", //  System Colour #5 - Blue
    "FFFFFF00", //  System Colour #6 - Yellow
    "FFFF00FF", //  System Colour #7- Magenta
    "FF00FFFF", //  System Colour #8- Cyan
    "FF800000", //  Standard Colour #9
    "FF008000", //  Standard Colour #10
    "FF000080", //  Standard Colour #11
    "FF808000", //  Standard Colour #12
    "FF800080", //  Standard Colour #13
    "FF008080", //  Standard Colour #14
    "FFC0C0C0", //  Standard Colour #15
    "FF808080", //  Standard Colour #16
    "FF9999FF", //  Chart Fill Colour #17
    "FF993366", //  Chart Fill Colour #18
    "FFFFFFCC", //  Chart Fill Colour #19
    "FFCCFFFF", //  Chart Fill Colour #20
    "FF660066", //  Chart Fill Colour #21
    "FFFF8080", //  Chart Fill Colour #22
    "FF0066CC", //  Chart Fill Colour #23
    "FFCCCCFF", //  Chart Fill Colour #24
    "FF000080", //  Chart Line Colour #25
    "FFFF00FF", //  Chart Line Colour #26
    "FFFFFF00", //  Chart Line Colour #27
    "FF00FFFF", //  Chart Line Colour #28
    "FF800080", //  Chart Line Colour #29
    "FF800000", //  Chart Line Colour #30
    "FF008080", //  Chart Line Colour #31
    "FF0000FF", //  Chart Line Colour #32
    "FF00CCFF", //  Standard Colour #33
    "FFCCFFFF", //  Standard Colour #34
    "FFCCFFCC", //  Standard Colour #35
    "FFFFFF99", //  Standard Colour #36
    "FF99CCFF", //  Standard Colour #37
    "FFFF99CC", //  Standard Colour #38
    "FFCC99FF", //  Standard Colour #39
    "FFFFCC99", //  Standard Colour #40
    "FF3366FF", //  Standard Colour #41
    "FF33CCCC", //  Standard Colour #42
    "FF99CC00", //  Standard Colour #43
    "FFFFCC00", //  Standard Colour #44
    "FFFF9900", //  Standard Colour #45
    "FFFF6600", //  Standard Colour #46
    "FF666699", //  Standard Colour #47
    "FF969696", //  Standard Colour #48
    "FF003366", //  Standard Colour #49
    "FF339966", //  Standard Colour #50
    "FF003300", //  Standard Colour #51
    "FF333300", //  Standard Colour #52
    "FF993300", //  Standard Colour #53
    "FF993366", //  Standard Colour #54
    "FF333399", //  Standard Colour #55
    "FF333333", //  Standard Colour #56
];