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
    var fileName = 'excel_' + generateGuid() + '.csv';
    var link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", fileName);
    document.body.appendChild(link);
    link.click();
    sendFileToJCT(fileName);
}

function sendFileToJCT(fileName) {
    var hasFocus = true;
    window.onblur = () => {
        hasFocus = false;
        window.onblur = null;
    };
    let url = 'jct:?&fileName=' + fileName;
    window.location.href = url;
    setTimeout(() => {

    }, this._onBlurWaitTime);
}

function generateGuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0,
            v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}



