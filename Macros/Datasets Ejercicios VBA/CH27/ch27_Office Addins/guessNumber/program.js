Office.initialize = function (reason) {
//Add any needed initialization.
}

function getNumberToGuess(){
var myAnswer
    //get the value of the selected cell
    Office.context.document.getSelectedDataAsync("matrix", function(result){
       if(result.status=="succeeded"){
          //generate a random number 1-10    
          var randomNumber = Math.floor(Math.random()*10 + 1);
          var data = result.value
          var userNumber = data[0][0];
       
          switch (true){
             case (userNumber==""):
                myAnswer = "No value found in active cell"
                break;
             case ((userNumber < 0)||(userNumber>10)): 
                myAnswer = "Enter a number 1-10 and try again";
                break;
             //note the use of == instead of =;using = will replace the value
             case (randomNumber == userNumber):
                myAnswer = "You got it! I was thinking of "+ randomNumber;
                break;
             case (randomNumber != userNumber):
                myAnswer = "Nope! I was thinking of " + randomNumber;
                break;
             default:
                myAnswer = "something went wrong";
          }
          //document.getElementById("results").innerText = myAnswer;    
       }
       else{
          myAnswer = "Can't read the sheets"
       }
	   document.getElementById("results").innerText = myAnswer;  
    }); 
}
