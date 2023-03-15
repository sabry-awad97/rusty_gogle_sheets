extern crate google_sheets4 as sheets4;
use hyper::client::HttpConnector;
use hyper_rustls::HttpsConnector;
use sheets4::api::{
    AddSheetRequest, BatchUpdateSpreadsheetRequest, CellData, GridRange, Request, Response, Sheet,
    SheetProperties, UpdateSheetPropertiesRequest, ValueRange,
};
use sheets4::api::{CellFormat, Color, RepeatCellRequest};
use sheets4::{hyper, hyper_rustls, Sheets};
use std::error::Error;

use crate::utils::{get_addr_int, get_cell_address};

type Service = Sheets<HttpsConnector<HttpConnector>>;

pub struct Spreadsheet {
    id: String,
    service: Service,
}

impl Spreadsheet {
    pub async fn new(id: String, service: Service) -> Result<Self, Box<dyn Error>> {
        Ok(Spreadsheet { id, service })
    }

    pub async fn get_worksheet_name(&self, id: i32) -> Result<Option<String>, Box<dyn Error>> {
        let request = self.service.spreadsheets().get(&self.id);
        let response = request.doit().await?.1;

        if let Some(sheets) = response.sheets {
            for sheet in sheets {
                if let Some(sheet_id) = sheet.properties.as_ref().and_then(|p| p.sheet_id) {
                    if sheet_id == id {
                        return Ok(sheet.properties.as_ref().and_then(|p| p.title.clone()));
                    }
                }
            }
        }

        Ok(None)
    }

    pub async fn get_worksheet_id(&self, title: &str) -> Result<Option<i32>, Box<dyn Error>> {
        let request = self.service.spreadsheets().get(&self.id);
        let response = request.doit().await?.1;

        if let Some(sheets) = response.sheets {
            for sheet in sheets {
                if let Some(sheet_title) = sheet
                    .properties
                    .as_ref()
                    .and_then(|p| p.title.as_ref().map(|t| t.to_string()))
                {
                    if sheet_title == title {
                        return Ok(sheet.properties.and_then(|p| p.sheet_id));
                    }
                }
            }
        }

        Ok(None)
    }

    pub async fn get_cell_values(&self, range: &str) -> Result<Vec<Vec<String>>, Box<dyn Error>> {
        let result = self
            .service
            .spreadsheets()
            .values_get(&self.id, &range)
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;

        Ok(result.1.values.unwrap_or_default())
    }

    pub async fn update_cells(
        &self,
        range: &str,
        values: Vec<Vec<String>>,
    ) -> Result<(), Box<dyn Error>> {
        let value_range = ValueRange {
            values: Some(values),
            ..Default::default()
        };
        let request = self
            .service
            .spreadsheets()
            .values_update(value_range, &self.id, range)
            .value_input_option("RAW");
        request
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;
        Ok(())
    }

    pub async fn insert_rows(
        &self,
        range: &str,
        values: Vec<Vec<String>>,
    ) -> Result<(), Box<dyn Error>> {
        let value_range = ValueRange {
            values: Some(values),
            ..Default::default()
        };
        let request = self
            .service
            .spreadsheets()
            .values_append(value_range, &self.id, range)
            .value_input_option("RAW")
            .insert_data_option("INSERT_ROWS");
        request
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;
        Ok(())
    }

    pub async fn write_column(
        &self,
        col: i32,
        start_row: i32,
        values: Vec<String>,
    ) -> Result<(), Box<dyn Error>> {
        let range = format!("{}", get_cell_address(start_row, col));
        let value_range = ValueRange {
            range: Some(range.to_owned()),
            values: Some(values.into_iter().map(|v| vec![v]).collect()),
            major_dimension: Some("COLUMNS".to_string()),
            ..Default::default()
        };
        let request = self
            .service
            .spreadsheets()
            .values_update(value_range, &self.id, &range)
            .value_input_option("RAW");
        request
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;
        Ok(())
    }

    pub async fn write_row(
        &self,
        row: i32,
        start_col: i32,
        values: Vec<String>,
    ) -> Result<(), Box<dyn Error>> {
        let range = format!(
            "{}{}",
            get_cell_address(row, start_col),
            get_cell_address(row, start_col + values.len() as i32 - 1)
        );
        let value_range = ValueRange {
            range: Some(range.to_owned()),
            values: Some(vec![values]),
            ..Default::default()
        };
        let request = self
            .service
            .spreadsheets()
            .values_update(value_range, &self.id, &range)
            .value_input_option("RAW");
        request
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;
        Ok(())
    }

    pub async fn rename_sheet(&self, sheet_name: &str) -> Result<(), Box<dyn Error>> {
        let requests = vec![Request {
            update_sheet_properties: Some(UpdateSheetPropertiesRequest {
                properties: Some(SheetProperties {
                    title: Some(sheet_name.to_string()),
                    ..Default::default()
                }),
                fields: Some("title".to_string()),
                ..Default::default()
            }),
            ..Default::default()
        }];

        let batch_update_spreadsheet_request = BatchUpdateSpreadsheetRequest {
            requests: Some(requests),
            ..Default::default()
        };

        let request = self
            .service
            .spreadsheets()
            .batch_update(batch_update_spreadsheet_request, &self.id);

        request
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;
        Ok(())
    }

    pub async fn create_sheet(&self, sheet_name: &str) -> Result<Option<i32>, Box<dyn Error>> {
        let mut properties = SheetProperties::default();
        properties.title = Some(sheet_name.to_string());

        let mut add_sheet_request = AddSheetRequest::default();
        add_sheet_request.properties = Some(properties);

        let mut single_request = Request::default();
        single_request.add_sheet = Some(add_sheet_request);

        let mut batch_update_spreadsheet_request = BatchUpdateSpreadsheetRequest::default();
        batch_update_spreadsheet_request.requests = Some(vec![single_request]);

        let request = self
            .service
            .spreadsheets()
            .batch_update(batch_update_spreadsheet_request, &self.id);

        let response = request
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?
            .1;

        let sheet_id = match response.replies {
            Some(reply) => {
                match reply.into_iter().find_map(|r| match r {
                    Response { add_sheet, .. } => Some(add_sheet),
                }) {
                    Some(add_sheet) => match add_sheet {
                        Some(add_sheet) => add_sheet.properties.and_then(|p| p.sheet_id),
                        _ => None,
                    },
                    None => None,
                }
            }
            None => None,
        };
        Ok(sheet_id)
    }

    pub async fn format_cell(
        &self,
        worksheet_name: &str,
        cell_address: &str,
        background_color: Option<Color>,
    ) -> Result<(), Box<dyn Error>> {
        // Get the ID of the worksheet.
        let worksheet_id = match self.get_worksheet_id(worksheet_name).await? {
            Some(id) => id,
            None => return Err("Worksheet not found".into()),
        };

        let (row, col) = get_addr_int(cell_address)?;

        // Define the range of cells to format
        let range = GridRange {
            sheet_id: Some(worksheet_id),
            end_row_index: Some(row),
            end_column_index: Some(col),
            ..Default::default()
        };

        // Define the formatting to apply to the cells
        let cell_format = CellFormat {
            background_color,
            ..Default::default()
        };

        let cell_data = CellData {
            user_entered_format: Some(cell_format),
            ..Default::default()
        };

        // Create the RepeatCellRequest to apply the formatting to the range of cells
        let repeat_cell_request = RepeatCellRequest {
            range: Some(range),
            cell: Some(cell_data),
            fields: Some("userEnteredFormat(backgroundColor)".to_string()), // Replace with the desired fields to update
        };

        let request = Request {
            repeat_cell: Some(repeat_cell_request),
            ..Default::default()
        };

        // Create the BatchUpdateSpreadsheetRequest to send the formatting request to the Google Sheets API
        let batch_update_request = BatchUpdateSpreadsheetRequest {
            requests: Some(vec![request]),
            ..Default::default()
        };

        // Send the request to the Sheets API.
        self.service
            .spreadsheets()
            .batch_update(batch_update_request, &self.id)
            .doit()
            .await
            .map_err(|e| format!("Error sending request: {:?}", e))?;

        Ok(())
    }
}
