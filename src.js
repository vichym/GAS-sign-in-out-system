function onSubmit(e){//Working 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openById('1T3gYnw-Mo9yUtV6PZlN1FgpjpgDYjUs2fxnWZ9oN9pE');
  var signoutSheet = ss.getSheetByName('Sign Out Students');
  var responsesheet = ss.getSheetByName('Form Responses');
  var signoutLastRow = signoutSheet.getLastRow();
  var lastResponseRow = responsesheet.getLastRow();
  var act = responsesheet.getRange(lastResponseRow, 2).getDisplayValue();
  var respondentEmail = responsesheet.getRange(lastResponseRow, 3).getDisplayValue();  
  var overnightList = ss.getSheetByName('Overnight Trip Permission'); 
  
  // the respondent is Signing out 
  if(act.equals('Signing Out')){
    
    //Check Overnight trip
    var overnightTrip = responsesheet.getRange(lastResponseRow, 12).getDisplayValue().equals('Yes');
    if(overnightTrip){ // overnight trip, check permission. 
      var overnightPermission = false; 

      for(var i=2; i <= overnightList.getLastRow(); i++){ // search permission List
        if(respondentEmail.equals(overnightList.getRange(i,2).getDisplayValue())){
          overnightPermission = true;
          overnightList.getRange(i, 3).setValue('Signed Out'); 
        }
      }
      
      if(overnightPermission){// process request
        // if there is no data in signout sheet
        if(signoutSheet.getLastRow()<=1){
          
          copySignoutData();//seeting value to Sign Out Sheet      
          borrowingPhone();// Borrowing phone and update phone list
          updateSignInList();}
        
        
        
        else{ // if there is data in signout sheet check if the respondent has already signout once
            var alreadySignedOut = false;
          
          // search for to check of the respondent already signed out once.  
          for(var i=2; i<= signoutLastRow;i++){
            if (signoutSheet.getRange(i, 3).getDisplayValue().equals(respondentEmail)){
              alreadySignedOut = true;} }
            
          if(alreadySignedOut){// if the respondent has already signout once
            var htmlBody = HtmlService.createHtmlOutputFromFile('sign_out_fail').getContent()
            GmailApp.sendEmail(respondentEmail, 'Signing Out Failed','',{ htmlBody: htmlBody })
            responsesheet.getRange(lastResponseRow, 13).setValue('Failed');
             responsesheet.getRange(lastResponseRow, 1, 1, 13).setFontColor('#cd0900');} 
          
          else if (!alreadySignedOut){ // if the respondent has no name on sign out list
            //seeting value to Sign Out Sheet
            copySignoutData();
            //ckeck borrowing phone
            borrowingPhone();
            updateSignInList();
          }// end of (!alreadySignOut)
        }
      } // end of Overnight Permission Okay 
      
      if(!overnightPermission){ // deny request and send email to respondent
        var htmlBody = HtmlService.createHtmlOutputFromFile('overnight_permission_deny').getContent()
            GmailApp.sendEmail(respondentEmail, 'Signing Out Failed','',{ htmlBody: htmlBody });
           responsesheet.getRange(lastResponseRow, 13).setValue('Failed');
          responsesheet.getRange(lastResponseRow, 1, 1, 13).setFontColor('#cd0900');
      }    
    }
    
//------------------------------------------------End of Overnight Trip-------------------------------------------------    
    
    if (!overnightTrip){ // not overnight Trip, no need to checnk permission, process the request
      
      // if there is no data in signout sheet
      if(signoutSheet.getLastRow()<=1){        
        copySignoutData(); //seeting value to Sign Out Sheet
        borrowingPhone();// Borrowing phone and update phone list 
        updateSignInList();
      }
      
      else{ // if there is data in signout sheet check if the respondent has already signout once      
        var alreadySignedOut = false;
        
        // search for to check of the respondent already signed out once. 
        for(var i=2; i<= signoutLastRow;i++){
          if (signoutSheet.getRange(i, 3).getDisplayValue().equals(respondentEmail)){
            alreadySignedOut = true;} 
        }
        
        if(alreadySignedOut){// if the respondent has already signout once
          var htmlBody = HtmlService.createHtmlOutputFromFile('sign_out_fail').getContent()
          GmailApp.sendEmail(respondentEmail, 'Signing Out Failed','',{ htmlBody: htmlBody }) ;
          responsesheet.getRange(lastResponseRow, 13).setValue('Failed');
          responsesheet.getRange(lastResponseRow, 1, 1, 13).setFontColor('#cd0900');} 
        
        else if (!alreadySignedOut){ // if the respondent has no name on sign out list
         
         copySignoutData(); //setting value to Sign Out Sheet
         borrowingPhone(); // update phone list
         updateSignInList();
        }// end of (!alreadySignOut)
      }
    }// end of Not Overnight trip
  } // end of if(act == signOut)
  

  // if the respondent is Signing In 
  if(act.equals('Signing In')){
    signIn();
    updateSignInList();
  }
} 

//------------------------------------------------------------------End of OnSubmit()---------------------------------------------------------------



function emailReminder(){
  var ss = SpreadsheetApp.getActive();
  var signoutSheet = ss.getSheetByName("Sign Out Students"); 
  var emailList = new Array(); 
  var extendedHours=2;
  var currentTime = new Date().getHours(); 
  var todayDate = new Date();
  
  
  // check the date of returning date and today date
  for (var i=2; i<= signoutSheet.getLastRow(); i++){
    // converting date into int in formate of yyyy, mm, dd to be compared the date of today date and returning date
    var returnDate = signoutSheet.getRange(i, 6).getValue();
    var sYyyy = Utilities.formatDate(new Date(returnDate), "GMT+9","yyyy");//  using Japanese time GMT+0900 (JST)
    var sMm = Utilities.formatDate(new Date(returnDate), "GMT+9","MM");
    var sDd = Utilities.formatDate(new Date(returnDate), "GMT+9","dd"); 
    var tYyyy = Utilities.formatDate(new Date(todayDate), "GMT+9","yyyy");
    var tMm = Utilities.formatDate(new Date(todayDate), "GMT+9","MM");
    var tDd = Utilities.formatDate(new Date(todayDate), "GMT+9","dd");
    
    if ((sYyyy + sMm + sDd) == (tYyyy + tMm + tDd)){ 
      // check if the expect return time exceed 2 hours limit
      var timestamp = signoutSheet.getRange(i, 7).getValue().getHours();
      var deadline = new Date(timestamp + extendedHours).getHours();
      
      if(deadline <= currentTime){
        // push the email of those who is already 2 hours later than 2 hours of expected return time
        emailList.push(signoutSheet.getRange(i,3).getDisplayValue());
        signoutSheet.getRange(i, 7).setFontColor('red'); 
      }
    }
  }
  
  if (emailList.length != 0){
    var empty = false;
    var htmlBody = HtmlService.createHtmlOutputFromFile('email_template').getContent()
    for(var i=0; i< emailList.length; i++){
      GmailApp.sendEmail( emailList[i], 'Missing Signing In','',{ htmlBody: htmlBody }) }
  }
}


//---------------------------------------------------------------End of emailReminder()----------------------------------------------------------------------------------

function clearResponse(){ // working
  var form = FormApp.openById('1T3gYnw-Mo9yUtV6PZlN1FgpjpgDYjUs2fxnWZ9oN9pE');
  form.deleteAllResponses(); // reset form response counter
  clearResponseSheet ();
  clearSignoutSheet(); 
}
  
function resetPhone(){
 var ss = SpreadsheetApp.getActive();
 var phoneNumSheet = ss.getSheetByName('School Phones');
  phoneNumSheet.getRange(phoneNumSheet.getActiveCell().getRow(), 2).setValue('Available').setFontColor('black');
  phoneNumSheet.getRange(phoneNumSheet.getActiveCell().getRow(), 3).clearContent();
  updatePhoneList();
}

function ManualSetBorrowerPhone(){
  var ss = SpreadsheetApp.getActive();
  var phoneNumSheet = ss.getSheetByName('School Phones');
  var activeRow = phoneNumSheet.getActiveCell().getRow();
  
  for(var i=2; i<=16;i++){
    if( !phoneNumSheet.getRange(i, 4).isBlank()){
      phoneNumSheet.getRange(i, 2).setValue('Unavailable').setFontColor('Red');
      phoneNumSheet.getRange(i, 3).setValue(phoneNumSheet.getRange(i, 4).getDisplayValue());
      phoneNumSheet.getRange(i, 4).clearContent();
    }}
  updatePhoneList();
}


function resetPhoneList(){
  var ss = SpreadsheetApp.getActive();
  var phoneNumSheet = ss.getSheetByName('School Phones');
  for(var i=2; i<=phoneNumSheet.getLastRow();i++){
    phoneNumSheet.getRange(i, 2).setValue('Available').setFontColor('black');
    phoneNumSheet.getRange(i, 3).clearContent();
  }
  updatePhoneList();
}

function updatePhoneList(){
  var form = FormApp.openById('1T3gYnw-Mo9yUtV6PZlN1FgpjpgDYjUs2fxnWZ9oN9pE');
  var ss = SpreadsheetApp.getActive();
  var phoneList = form.getItemById('2023521413').asMultipleChoiceItem();
  var phoneNumSheet = ss.getSheetByName('School Phones');
  var phoneAvailable = [];
  
  if(phoneNumSheet.getLastRow()<=1){
    phoneAvailable.push('All Phones are unavailable'); }
  else{
    for(var i = 2; i <= phoneNumSheet.getLastRow(); i++){
      if(phoneNumSheet.getRange(i, 2).getDisplayValue().equals('Available')){
        phoneAvailable.push(phoneNumSheet.getRange(i, 1).getDisplayValue());}
    } phoneList.setChoiceValues(phoneAvailable);
  }
}

function copySignoutData(){
  var form = FormApp.openById('1T3gYnw-Mo9yUtV6PZlN1FgpjpgDYjUs2fxnWZ9oN9pE');
  var ss = SpreadsheetApp.getActive();
  var responsesheet = ss.getSheetByName('Form Responses');
  var lastResponseRow = responsesheet.getLastRow();
  if(lastResponseRow>1){
    var signoutSheet = ss.getSheetByName('Sign Out Students');
    var signoutLastRow = signoutSheet.getLastRow();
    var signoutAfterLastRow = signoutLastRow+1; 

    signoutSheet.getRange(signoutAfterLastRow, 1).setValue(responsesheet.getRange(lastResponseRow, 1).getValue()); //    timestamp   11
    signoutSheet.getRange(signoutAfterLastRow, 3).setValue(responsesheet.getRange(lastResponseRow, 3).getDisplayValue()); // email 3-3
    signoutSheet.getRange(signoutAfterLastRow, 4).setValue(responsesheet.getRange(lastResponseRow, 4).getDisplayValue()); // Destination 4-4
    signoutSheet.getRange(signoutAfterLastRow, 10).setValue(responsesheet.getRange(lastResponseRow, 2).getValue()); // activity (act) 10-2
    signoutSheet.getRange(signoutAfterLastRow, 5).setValue(responsesheet.getRange(lastResponseRow, 5).getValue());// house 5-5
    signoutSheet.getRange(signoutAfterLastRow, 6).setValue(responsesheet.getRange(lastResponseRow, 6).getValue());//  return date 6-6
    signoutSheet.getRange(signoutAfterLastRow, 7).setValue(responsesheet.getRange(lastResponseRow, 7).getDisplayValue()); // return time 7-7
    signoutSheet.getRange(signoutAfterLastRow, 8).setValue(responsesheet.getRange(lastResponseRow, 8).getDisplayValue()); // return school phone number 8-8
    signoutSheet.getRange(signoutAfterLastRow, 9).setValue(responsesheet.getRange(lastResponseRow, 9).getDisplayValue()); // return personla phone number 9-9
    signoutSheet.getRange(signoutAfterLastRow, 11).setValue(responsesheet.getRange(lastResponseRow, 12).getDisplayValue()); // overnight trip
    var name = responsesheet.getRange(lastResponseRow, 3).getDisplayValue();
    var firstName = name.split(".")[1];
    var d = name.split(".")[2];// returns test
    var lastName = d.split("@")[0];
    var respondentName =firstName + " " + lastName;
    signoutSheet.getRange(signoutAfterLastRow, 2).setValue(respondentName);
    responsesheet.getRange(lastResponseRow, 13).setValue('Successful').setFontColor('#5cd302');
  }
}
          
function borrowingPhone(){
  var ss = SpreadsheetApp.getActive();
  var signoutSheet = ss.getSheetByName('Sign Out Students');
  var responsesheet = ss.getSheetByName('Form Responses');
  var lastResponseRow = responsesheet.getLastRow();
  var phoneId = responsesheet.getRange(lastResponseRow, 8);

  if(!phoneId.isBlank()){ 
    var borrowedPhone = phoneId.getDisplayValue(); 
    var phoneNumSheet = ss.getSheetByName('School Phones');
    for(var i = 2; i <= phoneNumSheet.getLastRow(); i++){
      var checkID = phoneNumSheet.getRange(i, 1).getDisplayValue();
      if(checkID.equals(borrowedPhone)){
        
        phoneNumSheet.getRange(i, 2).setValue('Unavailable').setFontColor('red'); // set Availability status to 'Unavailable in the 2nd column'
        
        var name = responsesheet.getRange(lastResponseRow, 3).getDisplayValue();
          var firstName = name.split(".")[1];
          var d = name.split(".")[2];// returns test
          var lastName = d.split("@")[0];
          var respondentName =firstName + " " + lastName;  
        phoneNumSheet.getRange(i, 3).setValue(respondentName).protect(); }
    }
  }
  updatePhoneList();
}


function signIn(){
  
  var ss = SpreadsheetApp.getActive();
  var signoutSheet = ss.getSheetByName('Sign Out Students');
  var overnightList = ss.getSheetByName('Overnight Trip Permission'); 
  var signoutLastRow = signoutSheet.getLastRow(); 
  var responsesheet = ss.getSheetByName('Form Responses');
  var lastResponseRow = responsesheet.getLastRow();
  var respondentEmail = responsesheet.getRange(lastResponseRow, 3).getDisplayValue();  
  
  for(var i=2; i<=signoutLastRow; i++){
    if (respondentEmail.equals(signoutSheet.getRange(i, 3).getDisplayValue())){
      signoutSheet.deleteRow(i);}
  }
  for(var i=2; i <= overnightList.getLastRow(); i++){ // search permission List
    if(respondentEmail.equals(overnightList.getRange(i,2).getDisplayValue())){
      overnightList.deleteRow(i); // delete respondent from Overnight Trip Sheet when sign in
    }
  }
}

function updateSignInList(){
  var ss = SpreadsheetApp.getActive();
  var signoutSheet = ss.getSheetByName('Sign Out Students');
  var form = FormApp.openById('1T3gYnw-Mo9yUtV6PZlN1FgpjpgDYjUs2fxnWZ9oN9pE');
  var namesList = form.getItemById('651380396').asMultipleChoiceItem();
  var studentNames = [];
  if(signoutSheet.getLastRow()>1){ // check if signout sheet is not empty, studentName array is to store name of those who sign out
    // grab the values in the first column of the sheet - use 2 to skip header row
    var namesValues= signoutSheet.getRange( 2, 2, signoutSheet.getLastRow()- 1).getDisplayValues();
    
    // there are people in sign out list
    for(var i = 0; i < namesValues.length; i++)
      // populate the drop-down with the array data
      if(namesValues[i][0] != ""){
        studentNames[i] = namesValues[i][0];
        namesList.setChoiceValues(studentNames);
      }
    namesList.setHelpText('Please check yourself in');
  }
  //if the sign out sheet is empty, set namelist to empty
  else{
    studentNames.push('Nobody Signs Out');
    namesList.setChoiceValues(studentNames); 
    namesList.setHelpText('Nobody signs out. You do not have to submit this form.');  
 }
}

function manualSignIn(){
 var ss = SpreadsheetApp.getActive();
 var signoutSheet = ss.getSheetByName('Sign Out Students');
 selectedRow = signoutSheet.getActiveCell().getRow();
  if(selectedRow!=1){
    signoutSheet.deleteRow(selectedRow); }
    updateSignInList();
}


function clearResponseSheet (){
  var ss = SpreadsheetApp.getActive();
  var responsesheet = ss.getSheetByName('Form Responses');
  if ((responsesheet.getLastRow()!=1)){
  responsesheet.deleteRows(2, responsesheet.getLastRow()-1);  // clear contents in form response sheet
  }
}

function clearSignoutSheet(){
  var ss = SpreadsheetApp.getActive();
  var signoutSheet = ss.getSheetByName('Sign Out Students');
  if ((signoutSheet.getLastRow()!=1)){
    signoutSheet.deleteRows(2, signoutSheet.getLastRow()-1) // clear content in sign out sheet. 
}
}












