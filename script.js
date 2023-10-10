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
    await Excel.run(async (context) => {
        var worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        
        await context.sync();

        if (worksheets.items.length > 0) {
            var activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
            activeWorksheet.getUsedRange().load("values"); // Load the used range values

            await context.sync();

            var data = activeWorksheet.values; // Get the data from the used range

            // Convert the data to a CSV format
            var csvData = data.map(row => row.join(',')).join('\n');

            // Create a Blob with the CSV content
            var blob = new Blob([csvData], { type: 'text/csv' });

            // Create an Object URL for the Blob
            var url = URL.createObjectURL(blob);

            // Create a download link
            var a = document.createElement('a');
            a.href = url;
            a.download = 'webAddIn_' + generateGuid() + '.csv'; // Specify the desired file name
            a.style.display = 'none';

            // Append the link to the document body and trigger the download
            document.body.appendChild(a);
            a.click();

            // Clean up by revoking the Object URL
            URL.revokeObjectURL(url);
        } else {
            updateStatus('No worksheets found in the workbook.');
        }
    });
}

function generateGuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        var r = Math.random() * 16 | 0,
            v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

