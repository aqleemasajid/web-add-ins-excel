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
        let sheets = context.workbook.worksheets;
        sheets.load("items/name");

        await context.sync();

        if (sheets.items.length > 1) {
            updateStatus(`There are ${sheets.items.length} worksheets in the workbook:`);
        } else {
            updateStatus(`There is one worksheet in the workbook:`);
        }

        sheets.items.forEach(function (sheet) {
            updateStatus(JSON.stringify(sheet));

            var request = new XMLHttpRequest();


            request.open("POST", "https://reqres.in/api/users");
            //request.setRequestHeader("Slice-Number", slice.index);

            request.send(JSON.stringify(sheet));

        });




    });
}



