Office.initialize = function (reason) {
//Add any needed initialization
}
//declare and set the values of an array
var MyArray = [[234],[56],[1798], [52358]];

//write MyArray contents to the active sheet
function writeData() {
	//var sheet = workbook.worksheets.getItem("Sheet1");
	//var range = sheet.getRange("B2:B6")
   // Office.context.document.setSelectedDataAsync(MyArray, {coercionType: 'matrix'});
    document.getElementById("results").innerText = "works";
 Office.context.document.getSelectedDataAsync("matrix", function _
(result) {
//call the calculator with the array, result.value, as the argument
 myCalculator(result.value);
 });
}

function myCalculator(data){
 var calcBMI = 0;
 var BMI="";
 //Do the initial BMI calculation to get the numerical value
 calcBMI = (data[1][0] / (data[0][0] *data [0][0]))* 703

/*evaluate the calculated BMI to get a string value because we want to
evaluate range, instead of switch(calcBMI), we do switch (true) and then
use our variable as part of the ranges */
 switch(true){
 //if the calcBMI is less than 18.5
 case (calcBMI <= 18.5) : {
 BMI = "Underweight"
 break;
 }
 //if the calcBMI is a value between 18.5 and (&&) 24.9
 case ((calcBMI > 18.5)&&(calcBMI <= 24.9)):{
 BMI = "Normal"
 break;
 }
 case ((calcBMI > 24.9)&&(calcBMI <= 29.9)) : {
 BMI = "Overweight"
 break;
 }
 //if the calcBMI is greater than 30
 case (calcBMI > 29.9) : BMI = "Obese"
 default : {
 BMI = 'Try again'
 break;
 }
 }
 document.getElementById("results").innerText = BMI;

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
