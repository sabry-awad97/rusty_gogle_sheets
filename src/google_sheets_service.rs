use std::error::Error;

use sheets4::oauth2::{read_service_account_key, ServiceAccountAuthenticator};
use sheets4::{hyper, hyper_rustls, Sheets};

use hyper::client::HttpConnector;
use hyper_rustls::HttpsConnector;
use sheets4::oauth2::authenticator::Authenticator;

use crate::spreadsheet::Spreadsheet;

type GoogleSheetsAuthenticator = Authenticator<HttpsConnector<HttpConnector>>;

pub struct GoogleSheetsService {
    auth: GoogleSheetsAuthenticator,
    client: hyper::Client<hyper_rustls::HttpsConnector<hyper::client::HttpConnector>>,
}

impl GoogleSheetsService {
    pub async fn new(key_path: String) -> Result<Self, Box<dyn Error>> {
        let service_account_key = read_service_account_key(key_path)
            .await
            .expect("service account key could not be read");

        let auth = ServiceAccountAuthenticator::builder(service_account_key)
            .build()
            .await
            .expect("failed to create authenticator");

        let client = hyper::Client::builder().build(
            hyper_rustls::HttpsConnectorBuilder::new()
                .with_native_roots()
                .https_or_http()
                .enable_http1()
                .enable_http2()
                .build(),
        );

        Ok(GoogleSheetsService { auth, client })
    }

    pub async fn connect_to_spreadsheet(&self, id: String) -> Result<Spreadsheet, Box<dyn Error>> {
        let service = Sheets::new(self.client.clone(), self.auth.clone());
        Spreadsheet::new(id, service).await
    }
}

