function purchasesIDFunction(date) {
	const SHEET_NAME = 'Metrika';
	const CONFIG = {
		SOURCE_COL: 1, // A
		DATE_COL: 2,   // B
		ID_COL: 3,     // C
		FORMULA_COL: 4 // D
	};
	
	try{
		const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
		const sheet = spreadsheet.getSheetByName(SHEET_NAME);
		
		if (!sheet) {
			throw new Error(`Sheet "${SHEET_NAME}" not found`);
		}
		if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
			throw new Error(`Invalid date format: ${date}. Expected YYYY-MM-DD`);
		}
		
		const dateExists = checkDateExists(sheet, date, CONFIG.DATE_COL);
    
		if (dateExists) {
			Logger.log(`Date ${date} already exists in column ${CONFIG.DATE_COL}`);
			return;
		}
		
		var purchaseIDs = metrikaDataFunction(date);
		if (!purchaseIDs.length) {
			Logger.log(`No purchase IDs found for date ${date}`);
			return;
		}
		
		const lastRow = sheet.getLastRow();
		const newData = purchaseIDs.map((id, i) => [
			"promopult_yandex_direct",               // SOURCE_COL
			date,                                    // DATE_COL
			id,                                      // ID_COL
			`=VLOOKUP(C${lastRow + i + 1}, Teplohod!A:C, 3, FALSE)` // FORMULA_COL
		]);
		
		sheet.getRange(lastRow + 1, 1, newData.length, newData[0].length)
			.setValues(newData);
      
		Logger.log(`Added ${purchaseIDs.length} records for date ${date}`);
		
	} catch (error) {
		Logger.log(`Error in purchasesIDFunction: ${error.message}`);
		throw error;
	}
}

function checkDateExists(sheet, targetDate, colNum) {
	const dateRange = sheet.getRange(1, colNum, sheet.getLastRow(), 1);
	const dates = dateRange.getValues();
  
	return dates.some(row => {
		const cellValue = row[0];
		if (cellValue instanceof Date) {
			return formatDate(cellValue) === targetDate;
		}
		return cellValue === targetDate;
	});
}

function formatDate(date) {
	return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}