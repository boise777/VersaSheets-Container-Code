/********************************* Build 3/11/2019 (rev 1) **********************************


 DESCRIPTION:
    
   VersaSheets Form Handling code located in Git Hub Repository 
     named "VersaSheets-Container-Code"

    
 AUTHOR:
    Steve Germain
    (208) 870-2727
    
 COPYRIGHT NOTICE:
   Copyright (c) 2009 - 2018 by Legatus, Inc.
   Copyright (c) 2019 by Stephen P. Germain
   
 NOTES:

   >>>>>> Change "XXXonOpen()" to "onOpen()" when used in host container, if necessary
   >>>>>> Update the Tools, Script, Resources Library Identifier to "VersaSheetsCommon" in host container


******************************************************************************/

var Version = "(v5.0)";
var SetupSheetName = 'Setup';

/******************************************************************************/
function onOpen() {
/******************************************************************************/
  var ReturnMessage = LoadGlobals(SetupSheetName);
  
  if (ReturnMessage != true){
    // Fatal if "ERROR" is detected  
    if (ReturnMessage.toUpperCase().indexOf("ERROR") > -1){
      Logger.log('*onOpen() ' + ReturnMessage);
      Browser.msgBox(ReturnMessage + ' Please contact the developer.');
      return ;
    }
  }
    
  Director("OnOpen", false);
} 

/******************************************************************************/
function updateFormResponse(e) {
/******************************************************************************/
  Director("updateFormResponse", true, e);
} 
  
/******************************************************************************/
function ProcessRequest_1() {
/******************************************************************************/
  Director("ProcessRequest_1", true);
}

/******************************************************************************/
function ProcessRequest_2() {
/******************************************************************************/
  Director("ProcessRequest_2", true);
}

/******************************************************************************/
function ProcessRequest_3() {
/******************************************************************************/
  Director("ProcessRequest_3", true);
}

/******************************************************************************/
function ScanForAlerts() {
/******************************************************************************/
  Director("ScanForAlerts", true);
}

/******************************************************************************/
function WeeklyScan() {
/******************************************************************************/
  Director("WeeklyScan", true);
}

/******************************************************************************/
function ResetGoogleForm(){
/******************************************************************************/
  Director("ResetGoogleForm", false);
}

/******************************************************************************/
function MoveSheetRows(){
/******************************************************************************/
  Director("MoveSheetRows", true);
}  

/******************************************************************************/
function DeleteSheetRows(){
/******************************************************************************/
  Director("DeleteSheetRows", false);
}  

/******************************************************************************/
function PerformCMGAudit(){
/******************************************************************************/
  Director("PerformCMGAudit", true);
}  

/******************************************************************************/
function EmergencyLockRelease(){
/******************************************************************************/
  var lock = LockService.getPublicLock();
  lock.releaseLock();
}

/******************************************************************************/
function CloseForms(){
/******************************************************************************/
  Director("CloseForms", true);
}  

/******************************************************************************/
function OpenForms(){
/******************************************************************************/
  Director("OpenForms", true);
}  

/******************************************************************************/
function RestoreRows(){
/******************************************************************************/
  Director("RestoreRows", false);
}  

/******************************************************************************/
/******************************************************************************/
/******************************************************************************/
function Director(CallingFunction, bNeedParams, e) {
/******************************************************************************/
/******************************************************************************/
/******************************************************************************/
  var func = '** Director ' + Version + ' - ';
  var Step = 100;
  Logger.log(func + Step + ' BEGIN Steps for "' + CallingFunction + '"');
  var bNoErrors = true;
  var title = CallingFunction + ' Procedures';
 
  /**************************************************************************/
  Step = 1000; // Load the oCommon object to the FormHandler function
  /**************************************************************************/
  Logger.log(func + Step + ' Build oCommon object');
  var oCommon = {};
  var ReturnMessage = null;
  
  oCommon = LoadCommon();
  
  if (oCommon.FatalErrorMessage != null){
    Step = 1100; 
    Browser.msgBox(oCommon.FatalErrorMessage);
    bNoErrors = false ;
  } else if(bNeedParams){
    Step = 1200; // Get the parameters to pass to the CallingFunction
    Logger.log(func + Step + ' Getting oMenu Parameters');  
    var oMenuParams ={};
    var bParamsFound = VersaSheetsCommon.GetMenuParams(oCommon, CallingFunction, oMenuParams);
    if (bParamsFound){
      title = oMenuParams["Menu Title"];
      oCommon.CallingMenuItem = title;
      //oMenuParams["Function Name"] = functionName;
    } else {
      Step = 1300; 
      //Browser.msgBox("Fatal Error 014: Unable to load required parameters. Please contact the developer");
      bNoErrors = false ;
    }
  }
  
  if (bNoErrors){
    Step = 1400; // Communicate with user
    var prog_message = 'Initializing...';
    VersaSheetsCommon.progressMsg(prog_message,title,-3);
    
    Step = 1500; // Capture / document any warnings that might have been found loading Globals
    if (oCommon.Globals['ErrorMessage'].toUpperCase().indexOf("WARNING") > -1){
      Logger.log(func + Step + ' (' + oCommon.Globals['ErrorMessage'] + ')');
      VersaSheetsCommon.LogEvent(oCommon.Globals['ErrorMessage'], oCommon);
      
      Step = 1510; // Clear the error message after it is posted
      var scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('ErrorMessage', '');
      
    }
      
    /**************************************************************************/
    Step = 2000; // Call and pass parameters to the CallingFunction
    /**************************************************************************/
    switch(CallingFunction) {
        
        /**************************************************************************/  
      case "OnOpen":
        /**************************************************************************/ 
        Step = 2010; // Execute OnOpen procedures
        oCommon.bSilentMode = true;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.onOpenProcedures(oCommon);
        
        
        try {
          Step = 2015; // Hide the Setup Tab
          oCommon.Sheets.getSheetByName(SetupSheetName).hideSheet();
          Step = 2016; // Hide the Event Messsages Tab, if it is used
          if (oCommon.EventMesssagesTab){
            oCommon.Sheets.getSheetByName(oCommon.EventMesssagesTab).hideSheet();  
          }
        }
        catch(err) {
          //  do nothing - probably a non-owner attempting to open the Sheet
          Logger.log(func + Step + ' Error: ' + err);
        }
        
        break;
        
        /**************************************************************************/  
      case "updateFormResponse":
        /**************************************************************************/ 
        Step = 2020; // Execute onFormSubmit procedures
        oCommon.bSilentMode = true;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);
  
        Step = 2022; // Get the triggerUid and Timestamp values from onSubmit event object "e"
                     // Ref: https://developers.google.com/apps-script/guides/triggers/events
        oCommon.FormTimestamp = e.namedValues['Timestamp'];
        oCommon.FormTriggerUid = e.triggerUid;
        Logger.log(func + Step + ' FormTimestamp: ' + oCommon.FormTimestamp 
           + ', FormTriggerUid: ' + oCommon.FormTriggerUid)
        
        if (VersaSheetsCommon.ParamCheck(oCommon.FormTimestamp)){
          VersaSheetsCommon.onFormSubmit(oCommon, oMenuParams);
        } else {
          Step = 2024; // No Timestamp value found, Log the ERROR Event and quit
          oCommon.DisplayMessage = ' ERROR 014 - Timestamp value not passed by onSubmit trigger event.';
          oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
          VersaSheetsCommon.LogEvent(oCommon.ReturnMessage, oCommon);
          Logger.log(oCommon.ReturnMessage);
        }
        
        break;
        
        /**************************************************************************/  
      case "ProcessRequest_1":
        /**************************************************************************/ 
        Step = 2030; // Execute procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.GenerateItems(oCommon, oMenuParams); 
        
        break;
        
        /**************************************************************************/  
      case "ProcessRequest_2":
        /**************************************************************************/ 
        Step = 2030; // Execute procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.GenerateItems(oCommon, oMenuParams);
        
        break;
        
        /**************************************************************************/  
      case "ProcessRequest_3":
        /**************************************************************************/ 
        Step = 2030; // Execute procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.GenerateItems(oCommon, oMenuParams);
        
        break;
        
        /**************************************************************************/  
      case "ScanForAlerts":
        /**************************************************************************/  
        Step = 2040; // Execute procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.AlertsScan(oCommon, oMenuParams); 
        
        break;
        
        /**************************************************************************/  
      case "WeeklyScan":
        /**************************************************************************/  
        Step = 2050; // Execute procedures
        oCommon.bSilentMode = true;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.AlertsScan(oCommon, oMenuParams);
        
        break;
        
        /**************************************************************************/  
      case "PerformCMGAudit":
        /**************************************************************************/  
        Step = 2060; // Execute procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        var CMGAuditTitle = '';
        VersaSheetsCommon.CMGAudit(CMGAuditTitle, oCommon, oMenuParams);
        
        break;
        
        /**************************************************************************/  
      case "ResetGoogleForm":
        /**************************************************************************/ 
        Step = 2070; // Execute OnOpen procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.updateFormLists(oCommon);
        
        break;
        
        /**************************************************************************/  
      case "MoveSheetRows":
        /**************************************************************************/ 
        Step = 2080; // Execute onFormSubmit procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.moveSelectedRows(oCommon, oMenuParams);
        
        break;
        
        /**************************************************************************/  
      case "DeleteSheetRows":
        /**************************************************************************/ 
        Step = 2090; // Execute onFormSubmit procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.deleteSelectedRows(oCommon);
        
        break;
        
        /**************************************************************************/  
      case "CloseForms":
        /**************************************************************************/ 
        Step = 2100; // Execute OpenCloseForms procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.OpenCloseForms(oCommon, oMenuParams);
        
        break;
         
        /**************************************************************************/  
      case "OpenForms":
        /**************************************************************************/ 
        Step = 2110; // Execute OpenCloseForms procedures
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);  
        
        VersaSheetsCommon.OpenCloseForms(oCommon, oMenuParams);
        
        break;
        
        /**************************************************************************/  
      case "RestoreRows":
        /**************************************************************************/ 
        Step = 2120; // Ree-execute onFormSubmit procedures using the Timestamp
                     //   taken from the form response data sheet
        oCommon.bSilentMode = false;
        Logger.log(func + Step + ' Executing "' + CallingFunction + '", SilentMode: ' 
                   + oCommon.bSilentMode);
        
        VersaSheetsCommon.RestoreRows(oCommon);
        
        break;
       
        /**************************************************************************/  
      default:
        /**************************************************************************/  
        Step = 3000; // Calling Function Not found.
        
        oCommon.ReturnMessage = func + Step + ' CallingFunction: "' + CallingFunction + '" Not found.';
        oCommon.DisplayMessage = 'Error 060 Encountered: CallingFunction: "' + CallingFunction + '" Not found.';
        VersaSheetsCommon.LogEvent(oCommon.ReturnMessage, oCommon);
        
        break;
    }  
  }
  /**************************************************************************/
  Step = 5000; // Display / record Success / Failure messages
  /**************************************************************************/
  Logger.log(func + Step + ' Return Message: ' + oCommon.ReturnMessage);
  Logger.log(func + Step + ' Display Message: ' + oCommon.DisplayMessage);
  Logger.log(func + Step + ' bSilentMode: ' + oCommon.bSilentMode);

  VersaSheetsCommon.WriteEventMessages("", oCommon);
  
  if (oCommon.DisplayMessage != ''){
    VersaSheetsCommon.progressMsg(oCommon.DisplayMessage,title,3);
    if(!oCommon.bSilentMode){
      Browser.msgBox(title, oCommon.DisplayMessage, Browser.Buttons.OK);
    }
  }
  
  VersaSheetsCommon.progressMsg("Bye, bye now!...",title,3);

  Step = 9999;
  Logger.log(func + Step + ' END');
  
  return;
  
}


/***************************************************************************************************************
****************************************************************************************************************
****************************************************************************************************************
*
*    Container-bound Functions Below
*
****************************************************************************************************************
****************************************************************************************************************
***************************************************************************************************************/

function RunTime(start) {
  // Useage:  RunTime = RunTime(start);
  var stop = new Date();
  var Runtime = Number(stop) - Number(start); // in milliseconds
  return Runtime;
}

function LoadGlobals(SetupSheetName) {
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked whenever the Globals have not been previously loaded
     
   USEAGE
     ReturnMessage = LoadGlobals(SetupSheetName);

   REVISION DATE:
    03-23-2018 - First Instance when discovered that Globals were not persisting in this container
     
   NOTES:

  ******************************************************************************/
  var func = "***LoadGlobals " + Version + " - ";
  Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  /******************************************************************************/
  Step = 1000; // Define the principal scalars and objects used everywhere
  /*******************************************************************************/
  var Globals = {};
  var oSourceSheets = SpreadsheetApp.getActiveSpreadsheet();
  var Globals = VersaSheetsCommon.BuildGlobals(oSourceSheets,SetupSheetName);
  
  /******************************************************************************/
  Step = 2000; //Set Persistent variable values for this container using the Properties Service
  /******************************************************************************/
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();  
  scriptProperties.setProperties(Globals);
  
  Logger.log(func + Step + Globals['ErrorMessage']);
  if (Globals['ErrorMessage']){
    Logger.log(func + Step + ' (' + Globals['ErrorMessage'] + ')');
    return Globals['ErrorMessage'];
  }
  
  return true;
 
}

function LoadCommon(e) {
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked whenever the Globals have not been previously loaded
     
   USEAGE
     var oCommon = {};
     var oCommon = LoadCommon();

   REVISION DATE:
    03-23-2018 - First Instance when discovered that Globals were not persisting in this container
     
   NOTES:

  ******************************************************************************/
  var func = "***LoadCommon " + Version + " - ";
  Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  // Get the Globals that were captured when onOpen was executed
  var scriptProperties = PropertiesService.getScriptProperties();
  var Globals = scriptProperties.getProperties();
  var Sheets = SpreadsheetApp.getActiveSpreadsheet(); 

  /******************************************************************************/
  Step = 1000; // Build the oCommon object
  /******************************************************************************/
  var oCommon = {};
  var oCommon = VersaSheetsCommon.BuildCommon(Sheets,Globals,e);
  // Verify oCommon
  //for(var Key in oCommon){
  //   Logger.log(func + Step + ' Key:' + Key + '  Value: ' + oCommon[Key]);
  //}
  // Verify Globals
  //for(var Key in oCommon.Globals){
  //   Logger.log(func + Step + ' Key:' + Key + '  Value: ' + oCommon.Globals[Key]);
  //}
  
  return oCommon;
  
}


/******************************************************************************/
function SummarizeTransactions() {
/******************************************************************************/
  var func = "*******SummarizeTransactions" + Version + " - ";
  var Step = 0;
  var title = 'Summarize Bank Transactions';
  Logger.log(func + Step + ' BEGIN');
  // Load the oCommon object to the FormHandler function
  var oCommon = {};
  oCommon = LoadCommon();
  if (oCommon.length <=0){ 
    Step = 100;
    var message = "ProcessTransactions procedure encountered oCommmon object error " + Version + ".";
    //VersaSheetsCommon.WriteEventMessages(message, oCommon);
    Logger.log(func + Step + message);
    VersaSheetsCommon.WriteEventMessages(message, oCommon);
    return; 
  }

  Step = 200;
  if(VersaSheetsCommon.SummarizeAccounts(oCommon)){
    var message = title + ' procedure completed successfully ' + Version + '.';
  } else {
    var message = title + ' procedure encountered errors and/or warnings ' + Version + '.';
  }
  Logger.log(func + Step + message);
  VersaSheetsCommon.WriteEventMessages(message, oCommon);
  VersaSheetsCommon.progressMsg(message,title,5);
  
  Step = 9999;
  Logger.log(func + Step + ' END');
  
}

/******************************************************************************/
function ProcessTransactions() {
/******************************************************************************/
  var func = "*******ProcessTransactions" + Version + " - ";
  var Step = 0;
  var title = 'Assign Accounts';
  Logger.log(func + Step + ' BEGIN');
  var oCommon = {};
  oCommon = LoadCommon();
  if (oCommon.length <=0){ 
    Step = 100;
    var message = "ProcessTransactions procedure encountered oCommmon object error " + Version + ".";
    //VersaSheetsCommon.WriteEventMessages(message, oCommon);
    Logger.log(func + Step + message);
    VersaSheetsCommon.WriteEventMessages(message, oCommon);
    return; 
  }

  Step = 200;
  if(VersaSheetsCommon.AssignAccounts(oCommon)){
    var message = title + ' procedure completed successfully ' + Version + '.';
  } else {
    var message = title + ' procedure encountered errors and/or warnings ' + Version + '.';
  }
  Logger.log(func + Step + message);
  VersaSheetsCommon.WriteEventMessages(message, oCommon);
  VersaSheetsCommon.progressMsg(message,title,5);
  
  Step = 9999;
  Logger.log(func + Step + ' END');
  
}

