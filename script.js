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
                // Handle the data (e.g., convert it to a downloadable format)
            });
    }).catch(function (error) {
        console.log(error);
    });
    
}


