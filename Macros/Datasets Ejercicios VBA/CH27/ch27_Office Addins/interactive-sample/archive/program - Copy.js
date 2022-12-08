Office.initialize = function (reason) {
//Add any needed initialization
}
//declare and set the values of an array
var MyArray = [[234],[56],[1798], [52358]];

//write MyArray contents to the active sheet
function writeData() {
    Office.context.document.setSelectedDataAsync(MyArray, {coercionType: 'matrix'});
}

/*reads the selected data from the active sheet
so that we have some content to read*/
function ReadData() {
    Office.context.document.getSelectedDataAsync("matrix", function (result) {
//if the cells are successfully read, print results in task pane
    if (result.status === "succeeded"){
            sumData(result.value);
        }
//if there was an error, print error in task pane
        else{
            document.getElementById("results").innerText = result.error.name;
        }
    });
}

/*the function that calculates and shows the result
in the task pane*/
function sumData(data) {
    var printOut = 0;

//sum together all the values in the selected range
    for (var x = 0 ; x < data.length; x++) {
        for (var y = 0; y < data[x].length; y++) {
            printOut += data[x][y];
        }
    }
//print results in task pane
   document.getElementById("results").innerText = printOut;
}
