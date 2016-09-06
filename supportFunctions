// easily authorize the script to run from the menu

function authorize(){
var respEmail = Session.getActiveUser().getEmail();
var formName = FormApp.getActiveForm().getTitle();
MailApp.sendEmail(respEmail,"Form Authorizer from "+formName, "Your form has now been authorized to send you emails");
}



//Set email, subject, and message body for form submission


function setEmailInfo(){
   
  var html = HtmlService.createHtmlOutputFromFile('EmailDialog')
        .setWidth(250)
      .setHeight(100);
  FormApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Email Notification Settings');
  
}


//click handler for setEmailInfo that saves values into script properties
function updateEmail(formObject){
  
  var email = formObject['authMail'];
  
  setProperty('EMAIL_ADDRESS',email);

}

//setting script properties
function setProperty(key,property){

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty(key);
  scriptProperties.setProperty(key,property);

}

//getting script properties
function getProperty(property){
  var scriptProperties = PropertiesService.getScriptProperties();
  var savedProp = scriptProperties.getProperty(property);
  return  savedProp;
}


function getWeekNumber() {
    /* Calculate Week Number */
    
    Date.prototype.getWeek = function() {
      var onejan = new Date(this.getFullYear(), 0, 1);
      return Math.ceil((((this - onejan) / 86400000) + onejan.getDay() + 1) / 7); 
    }
    
    var weekNumber = (new Date()).getWeek()-1;
    var today = (new Date()).getDay();
    
    // If Thu, Fri, Sat we are filling for next week
    // If Sun, Mon, Tue, Wed we are filling for this week
    if ( today > 3) weekNumber++;
    
  return weekNumber;
    
    /* End Week Number */
}
