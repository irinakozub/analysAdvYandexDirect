function metrikaDataFunction(date){
	const CONFIG = {
		counterId: "1", // Should be moved to script properties
		oauthToken: "XXX", // Should be moved to script properties
		apiBaseUrl: "https://api-metrika.yandex.net/stat/v1/data",
		utmSource: "utm_source"
	};
	
	try {
		// Validate input date format
		if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
			throw new Error(`Invalid date format: ${date}. Expected YYYY-MM-DD`);
		}

		const url = `${CONFIG.apiBaseUrl}?ids=${CONFIG.counterId}` +
			`&dimensions=ym:s:purchaseID` +
			`&metrics=ym:s:visits` +
			`&date1=${date}&date2=${date}` +
			`&filters=ym:s:UTMSource=='${CONFIG.utmSource}'`;
		  
		var options = {
			"async": true,
			"crossDomain": true,
			"method" : "GET",
			"headers" : {
			   "Authorization" : `OAuth ${CONFIG.oauthToken}`,
			   "Content-Type": "application/x-yametrika+json"
			},
			redirect: 'follow'
		};
		
		const response = UrlFetchApp.fetch(url, options);
		const responseData = JSON.parse(response.getContentText());
		
		if (!responseData.data || !Array.isArray(responseData.data)) {
			Logger.log("No purchase data found for date: " + date);
			return [];
		}
		
		const purchaseIDs = responseData.data
			.filter(item => item.dimensions && item.dimensions[0])
			.map(item => item.dimensions[0].name);

		Logger.log(`Found ${purchaseIDs.length} purchase IDs: ${purchaseIDs.join(", ")}`);
		return purchaseIDs;
	} catch (error) {
		Logger.log(`Error in metrikaDataFunction: ${error.message}`);
		throw error; // Re-throw for caller to handle
	}
}