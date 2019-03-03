const xlsx = require('node-xlsx').default;

// Parse a file
const workSheetsFromFile = xlsx.parse('../../../Downloads/Financial_Report.xlsx');

let output = {};

for (let a = 0; a < workSheetsFromFile.length; a++) {
	workSheetName = workSheetsFromFile[a].name;
	dataArray = workSheetsFromFile[a].data;
	if (workSheetName == "Document and Entity Information") {
		for (let b = 0; b < dataArray.length; b++) {
			if (dataArray[b][0] == 'Entity Registrant Name') output.companyName = dataArray[b][1];
			if (dataArray[b][0] == 'Document Period End Date') output.periodEndDate = dataArray[b][1];
			if (dataArray[b][0] == 'Document Fiscal Year Focus') output.fiscalYearFocus = dataArray[b][1];
		}
	}
	if (workSheetName == "CONSOLIDATED STATEMENTS OF INCO") {
		output.oldestTwelveMonthDate = dataArray[1][4];
		output.middleTwelveMonthDate = dataArray[1][3];
		output.newestTwelveMonthDate = dataArray[1][2];
		for (let b = 0; b < dataArray.length; b++) {
			if (dataArray[b][0] == 'Revenue') {
				output.oldestRevenue = dataArray[b][4];
				output.middleRevenue = dataArray[b][3];
				output.newestRevenue = dataArray[b][2];
			}
			if (dataArray[b][0] == 'Cost of revenue') {
				output.oldestCOGS = dataArray[b][4];
				output.middleCOGS = dataArray[b][3];
				output.newestCOGS = dataArray[b][2];
			}
		}
	}
}

console.log(JSON.stringify(output));