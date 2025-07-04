function analysAdvYandexDirect() {
	const DAYS_OFFSET = 1;
	const MILLISECONDS_IN_DAY = 24 * 60 * 60 * 1000;
	const SHEET_NAME = 'Promopult';
  
	try{
		const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
		const sheet = spreadsheet.getSheetByName(SHEET_NAME);
		if (!sheet) {
			throw new Error(`Sheet "${SHEET_NAME}" not found`);
		}

		const targetDate = new Date(Date.now() - DAYS_OFFSET * MILLISECONDS_IN_DAY);
		const formattedDate = targetDate.toISOString().split('T')[0]; // YYYY-MM-DD
		Logger.log(`Processing date: ${formattedDate}`);
		  
		var rangeDate = sheet.createTextFinder(date).findNext();
		if (!dateRange) {
			throw new Error(`Date ${formattedDate} not found in sheet`);
		}
		
		const rowNumber = dateRange.getRow();
		Logger.log(`Target row: ${rowNumber}`);
		
		var promopultData = promopultDataFunction(date);
		if (!promopultData) {
			throw new Error('No data received from promopultDataFunction');
		}
		
		sheet.getRange(`B${rowNumber}`).setValue(promopultData.get("shows"));
		sheet.getRange(`C${rowNumber}`).setValue(promopultData.get("clicks"));
		sheet.getRange(`N${rowNumber}`).setValue(promopultData.get("cost").replace(".",','));
		sheet.getRange(`D${rowNumber}`).setValue(promopultData.get("inside").replace(".",',')); 

		purchasesIDFunction(formattedDate);
	} catch (error) {
		Logger.log(`Error: ${error.message}`);
		throw error; 
	}
}
