extern crate google_sheets4 as sheets4;

mod google_sheets_service;
mod spreadsheet;
mod utils;

use std::env;
use std::error::Error;

use google_sheets_service::GoogleSheetsService;

use dotenv::dotenv;

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    dotenv().ok();
    let sheet_id = env::var("SPREADSHEET_ID").expect("SPREADSHEET_ID environment variable not set");
    let key_path = env::var("KEY_PATH").expect("KEY_PATH environment variable not set");
    let service = GoogleSheetsService::new(key_path).await?;
    let spreadsheet = service.connect_to_spreadsheet(sheet_id).await?;
    Ok(())
}
