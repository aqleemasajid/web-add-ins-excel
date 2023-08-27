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
            console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
        } else {
            console.log(`There is one worksheet in the workbook:`);
        }
    
        sheets.items.forEach(function (sheet) {
            console.log(sheet.name);
        });
    });
}

function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        } else {
            updateStatus(result.status);
        }
    });
}
function myEncodeBase64(docData)
{
    var s = "";
    for (var i = 0; i < docData.length; i++)
        s += String.fromCharCode(docData[i]);
    return window.btoa(s);
}
function sendSlice(slice, state) {
    var data = slice.data;

    if (data) {
        var fileData = myEncodeBase64(data);

        console.log(fileData);
 
        var request = new XMLHttpRequest();

        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                } else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "https://reqres.in/api/users");
        //request.setRequestHeader("Slice-Number", slice.index);

        request.send(data);
    }
}

function closeFile(state) {
    state.file.closeAsync(function (result) {

        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus("File closed.");
        } else {
            updateStatus("File couldn't be closed.");
        }
    });
}


///compare the excel script with ms excel script 