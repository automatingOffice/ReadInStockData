(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#submitStockData').click(getStockData)                
        });
    };

    function getStockData() {
        var stockId = $('#stockId').val(); 
        var dataUrl = 'https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=' + stockId + '&interval=5min&apikey=demo';
        console.log(dataUrl);
        $.ajax({
            type: 'GET',
            url: dataUrl,
            success: function (responseData) {
                var timeSeriesData = responseData['Time Series (5min)'];
                writeAPIDataToExcel(timeSeriesData);
            },
            error: function printError(errorMessage) {
                console.log(errorMessage);
            }
        });
    }

    function writeAPIDataToExcel(timeSeriesData) {
        Excel.run(function (ctx) {
            var excelDataArray = [];
            Object.keys(timeSeriesData).forEach(function (item) {
                var excelRowItem = [];

                excelRowItem[0] = item;
                excelRowItem[1] = timeSeriesData[item]['1. open'];
                excelRowItem[2] = timeSeriesData[item]['2. high'];
                excelRowItem[3] = timeSeriesData[item]['3. low'];
                excelRowItem[4] = timeSeriesData[item]['4. close'];
                excelRowItem[5] = timeSeriesData[item]['5. volume'];

                excelDataArray.push(excelRowItem);

            });

            var shareOutputRange = ctx.workbook.worksheets.getItem('ShareData').getRange('A1').getResizedRange(excelDataArray.length - 1, excelDataArray[0].length - 1).getOffsetRange(1, 0);
            var shareHeadingRange = ctx.workbook.worksheets.getItem('ShareData').getRange('A1:F1');
            shareHeadingRange.values = [['Date', 'Open', 'High', 'Low', 'Close', 'Volume']];
            shareHeadingRange.format.font.bold = true;
            shareOutputRange.values = excelDataArray;

            return ctx.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }



    

})();
