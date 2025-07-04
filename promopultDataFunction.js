function promopultDataFunction(date) {
	const CONFIG = {
		projectId: "1", // Should be moved to script properties
		apiToken: "xxx",  // Should be moved to script properties
		apiBaseUrl: 'https://api.promopult.ru/V0/project/ppc/'
	};
	
	
	try{
		if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
			throw new Error(`Invalid date format: ${date}. Expected YYYY-MM-DD`);
		}
		
		const url = `${CONFIG.apiBaseUrl}${CONFIG.projectId}/stat?dateFrom=${date}&dateTo=${date}`;
		Logger.log(url);
	
		var options = {
			"async": true,
			"crossDomain": true,
			"method" : "GET",
			"headers" : {
			   "X-Auth-Token" : CONFIG.apiToken,
			   "Content-Type": "application/json",
			   "cache-control": "no-cache"
			}
		};

		var response = UrlFetchApp.fetch(url, options);
		const responseText = response.getContentText();
		if (!responseText) {
			throw new Error("Empty response from API");
		}
		
		const responseTextReplaced = responseText.replace("\n",',');
		
		const metrics = responseTextReplaced.split(',');
		if (metrics.length < 10) { // Adjust based on actual API response
			throw new Error(`Unexpected response format: ${responseTextReplaced}`);
		}
		
		// Validate numeric values
		const parseMetric = (value, name) => {
			if (isNaN(parseFloat(value))) {
				throw new Error(`Invalid ${name} value: ${value}`);
			}
			return value;
		};
  
		return new Map([
			["cost", parseMetric(metrics[6], "cost")],
			["shows", parseMetric(metrics[7], "shows")],
			["clicks", parseMetric(metrics[8], "clicks")],
			["inside", parseMetric(metrics[9], "inside")]
		]);
		
	} catch (error) {
		Logger.log(`Error in promopultDataFunction: ${error.message}`);
		throw error; // Re-throw for caller to handle
	}
}