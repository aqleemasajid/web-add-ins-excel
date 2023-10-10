Office.initialize = function (reason) {

    $(document).ready(function () {
        $('#submit').click(function () {
            sendFile();
        });

        updateStatus("Ready to send file.");
    });
}

function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}

async function sendFile() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getUsedRange();
        range.load("values"); // Load the values within the range
        
        return context.sync()
            .then(function () {
                // Access the data
                var data = range.values;
                downloadCSV(data);
                // Handle the data (e.g., convert it to a downloadable format)
            });
    }).catch(function (error) {
        console.log(error);
    });
    
}

function downloadCSV(data) {
    var csvContent = "data:text/csv;charset=utf-8,";

    data.forEach(function (row) {
        var rowString = row.join(",");
        csvContent += rowString + "\n";
    });

    var encodedUri = encodeURI(csvContent);
    var link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "excel_data.csv");
    document.body.appendChild(link);
    link.click();
}


