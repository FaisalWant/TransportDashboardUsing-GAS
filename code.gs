
// Main SpreadSheet URL

var url="https://docs.google.com/spreadsheets/d/1FfpxMnyfdrzXDbd21KSfvKZT_4j7CJj3xvfBCW3cWfI/edit#gid=1060124669"




var Route = {};
Route.path = function(route, callback){
  Route[route]=callback;
}

//#################################

function doGet(e){ 

 // main function for ROUTE calling
 Logger.log("inside doGet function");
 Logger.log(e.parameters.v);

  Route.path("home", home);
  Route.path("search", loadSearch);
  Route.path("add",loadAdd );
  Route.path("reports",loadReports);
  Route.path("deployed", loaddeployed);
  Route.path("remote", loadRemote);
  Route.path("office", loadOffice);
  Route.path("rangreth", loadRangreth);
  Route.path("khayam", loadKhayam);
  Route.path("outstation", loadOutstation);
  Route.path("opt", loadOpt);
  // Route.path("categories", loadCategories);

  //--------------------------------------

   if (Route[e.parameters.v] ){  // if else case for detail having pagination
         
          
          if(e.parameters.id){  // we want to update a particular Stock

            //Storing the value as global
                var ss= SpreadsheetApp.openByUrl(url);
                var ws= ss.getSheetByName("global");
                ws.getRange("A"+1).setValue(e.parameters.id[0]);
             //=================================== 

            return loadAbout(e.parameters.id[0]);       
          }
            
        
            Logger.log("e.parameters");
            Logger.log(e.parameters);   
            return Route[e.parameters.v]();


    }
   else{
     return home();
     
     
    //  HtmlService.createTemplateFromFile("index").evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

   }
 
}// end do get



// index view of the Dashboard

// Load Functions 
function home(){


      var temp=HtmlService.createTemplateFromFile("index");
      var dataStatus=getStatusOccurance();
      var optStatus= getOptStatus()
      var locData=LocationStatusCount()


      temp.dataStatus=dataStatus;
      temp.optStatus= optStatus;
      temp.locData=locData
      return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}

//#################################
function loadOpt(){

   var temp=HtmlService.createTemplateFromFile("optout");
      return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}




//============================

function loadSearch(){
      var url= ScriptApp.getService().getUrl(); // passing current app url to the page
      var temp=HtmlService.createTemplateFromFile("search");
      var backUrl=url+"?v=home"
      temp.url=backUrl
      return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  



}
//##################################
function loadAdd(){
   var temp=HtmlService.createTemplateFromFile("add");
      return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}


function loadReports(){


 var temp=HtmlService.createTemplateFromFile("reports");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);



}

function loaddeployed(){
var temp=HtmlService.createTemplateFromFile("deployed");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}

function loadRemote(){
var temp=HtmlService.createTemplateFromFile("remote");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}

function loadOffice(){
var temp=HtmlService.createTemplateFromFile("office");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}

function loadRangreth(){
var temp=HtmlService.createTemplateFromFile("rangreth");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}

function loadKhayam(){
var temp=HtmlService.createTemplateFromFile("khayam");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}
function loadOutstation(){
var temp=HtmlService.createTemplateFromFile("outstation");
return temp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}




// Load Functions End Here



//##################################   Utility Functions 



function getOptOut(){


  var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=opt"
  Logger.log("addUrl", addUrl);
  return addUrl



}

function getRangreth(){
   var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=rangreth"
  Logger.log("addUrl", addUrl);
  return addUrl
}

function getKhayam(){
   var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=khayam"
  Logger.log("addUrl", addUrl);
  return addUrl
}


function getOutstation(){
   var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=outstation"
  Logger.log("addUrl", addUrl);
  return addUrl
}





function getOffice(){
   var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=office"
  Logger.log("addUrl", addUrl);
  return addUrl
}

function getRemote(){
   var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=remote"
  Logger.log("addUrl", addUrl);
  return addUrl
}

function getdeployed(){
   var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=deployed"
  Logger.log("addUrl", addUrl);
  return addUrl
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getHomeUrl(){
     var url= ScriptApp.getService().getUrl();
     var homeUrl= url+"?v=home";
     return homeUrl
}



function getSearch(){

  var url= ScriptApp.getService().getUrl();
  var searchUrl= url+"?v=search"
  Logger.log("searchUrl", searchUrl);
  return searchUrl
}


function getAdd(){

  var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=add"
  Logger.log("addUrl", addUrl);
  return addUrl
}


function getReport(){


  var url= ScriptApp.getService().getUrl();
  var addUrl= url+"?v=reports"
  Logger.log("addUrl", addUrl);
  return addUrl


}


// Utility Functions End Here



//============================================== BackEnd Functions 

function uuid() {  // unique ID function
         
          var uuid_array = [];
          var ss= SpreadsheetApp.openByUrl(url);
          var dataSheet = ss.getSheetByName("Data");
          var getLastRow = dataSheet.getLastRow();
          if(getLastRow > 1) {
                  var uuid_values = dataSheet.getRange(2, 1, getLastRow - 1, 1).getValues(); 
                  for(i = 0; i < uuid_values.length; i++)
                          {
                            uuid_array.push(uuid_values[i][0]);
                          }
                  var x_count = 0;
                  do {
                  var y = 'false';
                  var uuid_value = Utilities.getUuid(); 

                  if(uuid_array.indexOf(uuid_value) == -1.0)
                        {
                          y = 'true';
                          Logger.log(uuid_value);
                          return uuid_value;   
                        } 
                  x_count++;
                  } while (y == 'false' && x_count < 5);
                } 
          else {
                
                 return Utilities.getUuid();
          }
}
//-------------------------------------------------------


function AddRecord(userInfo) {

  var uniqueID = uuid();
  var found_record = false;
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Data");
  var getLastRow = dataSheet.getLastRow();
  
  Logger.log("getLastRow");
  //Logger.log(getLastRow);
  
  var date= Utilities.formatDate(new Date(userInfo.date), "IST", "yyyy-MM-dd");
  // var ostime=
  // var oetime= 
  
  Logger.log("userInfo.oetime----");
  //Logger.log(userInfo.oetime);
  
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue() == '')
    {
      dataSheet.getRange('A' + i + ':M' + i).setValues([[uniqueID, userInfo.fname, userInfo.lname,userInfo.dept,userInfo.address,userInfo.emStat,userInfo.offLoc,date,userInfo.cstatus,userInfo.optout,userInfo.ostime,userInfo.oetime. userInfo.pnum]]);

      found_record = true;
      break;
    }
  }
  if(found_record == false)

  { 
        dataSheet.getRange('A' + i + ':M' + i).setValues([[uniqueID, userInfo.fname, userInfo.lname,userInfo.dept, userInfo.address ,userInfo.emStat,userInfo.offLoc,date,userInfo.cstatus,userInfo.optout,userInfo.ostime,userInfo.oetime, userInfo.pnum
      
    ]]);
  }
  return 'SUCCESS';
  
}






//-------------------------------------------------------
function getRecords() { 


        var return_Array = [];
        var ss= SpreadsheetApp.openByUrl(url);
        var dataSheet = ss.getSheetByName("Data"); 
        var getLastRow = dataSheet.getLastRow();
       // Logger.log(getLastRow)


        for(i = 2; i <= getLastRow; i++)
                {
                  if(dataSheet.getRange(i, 1).getValue() != '')  //(row, col)(2,1)
                  {  
                    var timezone = Session.getScriptTimeZone()
                    var format = 'hh:mm a' // Set date format for output
                    
                    // normalizing the office start time-----

                    if (dataSheet.getRange(i, 11).getValue() !=""){

                        var newDate=dataSheet.getRange(i, 11).getValue();
                    }
                 

                    // normalizing the office end time ------
                   if (dataSheet.getRange(i, 12).getValue() !="") {

                          var newDate2=dataSheet.getRange(i, 12).getValue()


                   }

                    


                    //==================================
                            return_Array.push([
                            dataSheet.getRange(i, 1).getValue(), 
                            dataSheet.getRange(i, 2).getValue(),
                            dataSheet.getRange(i, 3).getValue(),
                            dataSheet.getRange(i, 4).getValue(),
                            dataSheet.getRange(i, 5).getValue(),
                            dataSheet.getRange(i, 6).getValue(), 
                            dataSheet.getRange(i, 7).getValue(), 
                            Utilities.formatDate(new Date(dataSheet.getRange(i, 8).getValue()),"IST", "yyyy-MM-dd"),
                            dataSheet.getRange(i, 9).getValue(),
                            dataSheet.getRange(i, 10).getValue(),

                            dataSheet.getRange(i, 11).getValue() !=""? Utilities.formatDate(newDate, timezone, format):"N/A",
                            dataSheet.getRange(i, 12).getValue() !=""? Utilities.formatDate(newDate2,timezone, format):"N/A",
                            dataSheet.getRange(i, 13).getValue() 
                                                                 
                            
                            
                           
                           // Utilities.formatDate(newDate, timezone, format)

                            // dataSheet.getRange(i, 12).getValue(),
                            // dataSheet.getRange(i, 13).getValue(),
                            // dataSheet.getRange(i, 14).getValue(),
                            // dataSheet.getRange(i, 15).getValue(),
                            // dataSheet.getRange(i, 16).getValue(),
                           // Utilities.formatDate(new Date(dataSheet.getRange(i, 17).getValue()), "IST", "MM-dd-yyyy"),
                           // Utilities.formatDate(new Date(dataSheet.getRange(i, 11).getValue()), "IST", 'HH:mm:ss Z')
                    
                    ]);
                  }
                }  
        // return return_Array;  
   //Logger.log("return Array")
   //Logger.log(return_Array);
  return return_Array;
}
//-------------------------------------------------------



function searchRecords(searchObject)
 {

              
              var name=searchObject.name;
              var department=searchObject.department;
              var officeLocation=searchObject.officeLocation;
              var empStatus=searchObject.empStatus;          
              var date=searchObject.date;
              var opt=searchObject.opt

              // var name="Aatif";
              // var department="Professional Services ";
              // var officeLocation="Corporate- Rangreth";
              // var empStatus="FTE";          
              // var date="";

              


              var pageId=searchObject.pageId||1;
              Logger.log("pageId is given by---------")
              Logger.log(pageId);
               
            //  if (searchObject.pagetId !=0){
            //   pageId=
               
            //    Logger.log("inside pageId----------");
            //    Logger.log(pageId);
            //  }
            


              //Logger.log("searchobject");
              //Logger.log(searchObject);

              var mydata =[];
              var allRecords = getRecords();
              
              //Logger.log("allRecords");
              //Logger.log(allRecords);
              
              
              allRecords.forEach(function(value, index) {

                var evalRows = [];  // is flushed with everyiteration



                if(name != "")
                {
                  if(value[1].toUpperCase() == name.toUpperCase())
                  { 
                      
                    evalRows.push('true');
                   // Logger.log(evalRows);
                  } else {
                    evalRows.push('false');
                  }
                }
                else
                {
                  evalRows.push('false');
                }


                if(department != "")
                {
                  if(value[3].toUpperCase() == department.toUpperCase()) 
                  {
                    evalRows.push('true');
                    //Logger.log(evalRows);
                  } else {
                    evalRows.push('false');
                  }
                }
                else
                {
                  evalRows.push('false');
                }

                if(empStatus != "")
                {
                  if(value[8].toUpperCase() == empStatus.toUpperCase())
                    {
                    evalRows.push('true');
                    // Logger.log(evalRows);
                  } else {
                    evalRows.push('false');
                  }
                }
                else
                {
                  evalRows.push('false');
                }



            if(officeLocation != "")
            {
                if(value[6].toUpperCase() == officeLocation.toUpperCase()) {
                  evalRows.push('true');
                  //Logger.log(evalRows);
                } else {
                  evalRows.push('false');
                }
            }
            else
            {
                evalRows.push('false');
            }


              // ------------------------------------for further fields search

            if(date != '')
            {
                if(value[7] == date) {
                  evalRows.push('true');
                } else {
                  evalRows.push('false');
                }
            }
            else
            {
                evalRows.push('false');
            }



            if(opt != '')
            {
                if(value[9].toUpperCase() == "YES") {
                  evalRows.push('true');
                } else {
                  evalRows.push('false');
                }
            }
            else
            {
                evalRows.push('false');
            }

                // if(email != '')
                // {
                //    if(value[7].toUpperCase() == email.toUpperCase()) {
                //      evalRows.push('true');
                //    } else {
                //      evalRows.push('false');
                //    }
                // }
                // else
                // {
                //    evalRows.push('true');
                // }

                Logger.log('Final Eval Array------------');
                Logger.log(evalRows);

                if(evalRows.indexOf("true") > -1)
                {
                  mydata.push(value);    
                  
                }

                // Logger.log('returnRows------------');
                // Logger.log(returnRows);

                // one iteration ends here

              });  // forEach ends here

            //  Logger.log('returnRows------------');
            //  Logger.log(mydata);
              
            //  returnRows.forEach(function(value,index){
            //    Logger.log(value);
            //    Logger.log("\n")
            //  });

            //var mydatas=[["fcb82067-9619-423e-a160-09dbd8be065b", "LT-0105", "Rajab jamal" , "Deployed", "keyboard", "HP", "RTLB723DE", "256 SSD", "8.0", "admin@iquasar.com", "Sat Jul 24 00:00:00 GMT+05:30 2021", "54321.0", "Unavailable", "Body Replacement needed///", "Rangreth", "Absar"],["9095fddf-4833-4420-b4db-a7a8bb4a8292qr", "LT-0001", "faisal Afzal" , "Pending", "Keyboard"," Dell", "Vostro", "512 SSD", "8.0", "admin@iquasar.com", "Tue Jul 27 00:00:00 GMT+05:30 2021", "12345.0", "Unavailable"," issue has been resolved", "Rangreth", "Absar"], ["7dd9e6bd-a671-4773-ad5d-c27a798345d9mk", "LT-0003123", "MotherLand High School", "Redy to Deploy", "Keyboard", "Dell", "Vostro", "500.0", "8.0", "admin@iquasar.com", "Thu Jul 29 00:00:00 GMT+05:30 2021", "12345.0", "Unavailable", "School received this device with any slip/", "Zirakpur", "Absar"], ["fae53131-a1fb-46c2-8cdb-a90683ae135f", "LT0123Rer", "rameez jan", "Deployed", "Keyboard", , , , , , "Fri Aug 20 00:00:00 GMT+05:30 2021", "12345.0", "Unavailable", "device is good", "Zirakpur"," Absar"]];


            // Data is returned only when JSON.stringify is used
            //  Logger.log("Json.stringify------");
            //Logger.log(JSON.stringify(mydata));
            // return (JSON.stringify(mydata));
            //Logger.log("mydata returned") 
            //Logger.log(mydata)  
            
            
            //return JSON.stringify(mydata);
              
            return test1000(mydata,pageId);// paginated Version
}

//============================================================================





function ResetAll(searchObject)
{

              // var name=searchObject.name;
              // var department=searchObject.department;
              // var officeLocation=searchObject.officeLocation;
              // var empStatus=searchObject.empStatus;          
              // var date=searchObject.date;

              
              var name="";
              var department="ProfessionalServices";
              var officeLocation=""
              var empStatus=""          
              var date=""
              var return_Array = [];

              var ss= SpreadsheetApp.openByUrl(url);
              var dataSheet = ss.getSheetByName("Data"); 
              var getLastRow = dataSheet.getLastRow();



        for(i = 2; i <= getLastRow; i++)
                {
                  if(dataSheet.getRange(i, 1).getValue() != '')  //(row, col)(2,1)
                  {  

                    if( dataSheet.getRange(i, 4).getValue() == department)   // checking for department
                      {
                          
                            dataSheet.getRange(i, 11).setValue(""),
                            dataSheet.getRange(i, 12).setValue("")
                      }



                    //==================================
                          
            
                  }
                }  


  
}

// function ends here







//===========================================================================


function UpdateRecord(userInfo){


  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Data"); 

  var getLastRow = dataSheet.getLastRow();
  var table_values = dataSheet.getRange(2, 1, getLastRow - 1, 15).getValues();


  for(i = 0; i < table_values.length; i++)
  {
    if(table_values[i][0] == userInfo.record_id)    // checking the primary key column
    {
      dataSheet.getRange(i+2, 2).setValue(userInfo.name);
      dataSheet.getRange(i+2, 3).setValue(userInfo.lname);
      dataSheet.getRange(i+2, 4).setValue(userInfo.department);
      //dataSheet.getRange(i+2, 5).setValue();

      dataSheet.getRange(i+2, 6).setValue(userInfo.empStat);
      dataSheet.getRange(i+2, 7).setValue(userInfo.location);


      dataSheet.getRange(i+2, 8).setValue(Utilities.formatDate(new Date(userInfo.date), "IST", "yyyy-MM-dd"));
      //dataSheet.getRange(i+2, 11).setValue(userInfo.date);

      dataSheet.getRange(i+2, 9).setValue(userInfo.currentStat);

      dataSheet.getRange(i+2, 10).setValue(userInfo.optStatus);
      Logger.log("userInfo.officeStart-------------------")
      //Logger.log(userInfo.officeStart)

      Logger.log("userInfo.officeStart-------------------")
      //Logger.log(userInfo.officeEnd)
      dataSheet.getRange(i+2, 11).setValue(userInfo.officeStart);
      dataSheet.getRange(i+2, 12).setValue(userInfo.officeEnd);
      dataSheet.getRange(i+2, 13).setValue(userInfo.phno)
      
         

      // add more fields

    } // if ends



}// for ends here

return 'SUCCESS';

}
//=================================================================

function DeleteRecord(key)
{
 
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Data");
  var getLastRow = dataSheet.getLastRow();
  var table_values = dataSheet.getRange(2, 1, getLastRow - 1, 15).getValues();
  for(i = 0; i < table_values.length; i++)
  {
    if(table_values[i][0] == key)
    {
      var rowNumber = i+2;
      dataSheet.getRange('A' + rowNumber +':M' + rowNumber).clearContent();
      
    }   
  }
  return 'SUCCESS';
}

//====================================================================

function test1000(searchedData,pageId){   // search function shld return through this function 
  Logger.log("pageId");
  Logger.log(pageId);
  return getPaginated(searchedData,pageId);
}
    

// paginator function

function getPaginated(searchedData,page){
    //var ss= SpreadsheetApp.openByUrl(url);
    //var dataSheet = ss.getSheetByName("Data");
    //var rows=dataSheet.getDataRange().getValues();
    //Logger.log("ROWS--------------");
    //Logger.log(rows);
    var posts=searchedData
    Logger.log("Posts is the whole data");
   // Logger.log(posts);
    var pg=page;
    var maxPosts=10;
    var json=JSON.stringify(paginator(posts, pg, maxPosts));
    return json;
}

      
//++++++++++++++++++++++++++++++++++++++++

      
  
      
//+++++++++++++++++++++++++++++++++++++++++++++++++
      
      
function paginator(posts, page, maxPosts){
  page=parseInt(page);
  var offset=(page-1)* maxPosts;
  var items= posts.slice(offset,(offset+maxPosts));   // item here is the subset of data that we are getting
  var totalPages=Math.ceil(posts.length/maxPosts);
  var nextPage=(totalPages > page) ? page+1 : null;
  var prePage= page -1 ? page -1 : null;
  var obj={

    page:page,
    maxPP:maxPosts,
    prev:prePage,
    next:nextPage,
    totalItems:items.length,
    totalPages:totalPages,
    data:items
  }
 // Logger.log(obj.data);
  return obj;


}
//---------------------------------
function newSheet(){

  // Initializing the Reports Sheet first  
   //const id= "1FfpxMnyfdrzXDbd21KSfvKZT_4j7CJj3xvfBCW3cWfI"
   //const ss= SpreadsheetApp.openById(id);
   const ss= SpreadsheetApp.openByUrl(url);
   
  const sheets= ss.getSheets()

  //  Logger.log(sheets)
  const newName= "Report"
  var  sheet= ss.getSheetByName("Report");
   
   if (sheet == null){

     sheet = ss.insertSheet();
     sheet.setName(newName);
   }
    


}

//=============================================

function tConvert (time) {
  //var time= "09:30"
  // Check correct time format and split into components
  time = time.toString ().match (/^([01]\d|2[0-3])(:)([0-5]\d)(:[0-5]\d)?$/) || [time];
  
  if (time.length > 1) { // If time format correct
    time = time.slice (1);  // Remove full string match value
   
    time[5] = time[0] < 12 ? ' AM' : ' PM'; // Set AM/PM

    time[0] = time[0] % 12 || 12; // Adjust hours
    if (time[0] < 10)
    {
      time[0]="0"+time[0]
    }
   
  }

  return time.join (''); // return adjusted time or original string
}


//----------------------------------------------------

/**
 * creating new spreadsheet using appscript
 * 
 * 
 * searchObject
 */

function makeNewSheet(searchObject){

  // Initializing the Reports Sheet first  
  //  const id= "1FfpxMnyfdrzXDbd21KSfvKZT_4j7CJj3xvfBCW3cWfI"
  //  const ss= SpreadsheetApp.openById(id);
      const ss= SpreadsheetApp.openByUrl(url);
  //  const sheets= ss.getSheets()

  //  Logger.log(sheets)
  //  const newName= "Report"
  var  sheet= ss.getSheetByName("Report");
  var getLastRow = sheet.getLastRow();

  for(i = 1; i <= getLastRow; i++){
  var range = sheet.getRange("A"+i+":L"+i);
  range.clearContent();
  }

// Search from parent sheet 
  //var department="ProfessionalServices"
  var mydata =[];



  var allRecords = getRecords();
  

  allRecords.forEach(function(value, index) {

    var evalRows = []; 

    // Logger.log(value[7]);
     Logger.log(value[10]);
    //   Logger.log(value[11]);

    Logger.log("------------------Input Values===================");

    //Logger.log(time);

    var val10= value[10];

    var val11= value[11];

    var ofstime=tConvert(searchObject.ostime)
    var ofetime=tConvert(searchObject.oetime)


    // Logger.log(val10)
    // Logger.log(val11)

    // Logger.log("input time is ----")
    Logger.log("input Date Range")
    // Logger.log(searchObject.date)
    // Logger.log(searchObject.date2)



    Logger.log("=======================================");
//===============================


//============================
    if((searchObject.date != "") && (ofstime != "") && (ofetime != ""))
    {   var date3=new Date(value[7])

      if( (date3 >= new Date (searchObject.date)) && ( date3 <= new Date (searchObject.date2)) && ( val10 == ofstime) && ( val11 == ofetime)  )
      { 
          
        evalRows.push('true');
        Logger.log(evalRows);
      } else {
        evalRows.push('false');
      }
    }
    else if ((searchObject.date != "") && (ofstime == "") && (ofetime == "") ) //Date Range given but not time
    {
      var date3=new Date(value[7])
      if( (date3 >= new Date (searchObject.date)) && ( date3 <= new Date (searchObject.date2))){

            evalRows.push('true');
             Logger.log(evalRows);

      }
      
    
    }

    else if ((searchObject.date != "") && (ofstime != "") && (ofetime == "") ) // Date Range is given but no end time
    {
      var date3=new Date(value[7])
      if( (date3 >= new Date (searchObject.date)) && ( date3 <= new Date (searchObject.date2) && ( val10 == ofstime))){

            evalRows.push('true');
             Logger.log(evalRows);

      }
      
    
    }
  

   else if ((searchObject.date != "") && (ofstime == "") && (ofetime != "") ) // Date Range is given but not start time
    {
      var date3=new Date(value[7])
      if( (date3 >= new Date (searchObject.date)) && ( date3 <= new Date (searchObject.date2) && ( val11 == ofetime))){

            evalRows.push('true');
             Logger.log(evalRows);

      }
      
    
    }





   
    else {
      evalRows.push('false');

    }

    if(evalRows.indexOf("true") > -1)
      {
        mydata.push(value);    
        
      }



  }); // for ends here


 

// Add data into new sheet-------------
   
for(i = 1; i <= mydata.length; i++){

    sheet.getRange('A' + i + ':L' + i).setValues([[ mydata[i-1][1], mydata[i-1][2],mydata[i-1][3],mydata[i-1][4],mydata[i-1][5],mydata[i-1][6],mydata[i-1][7],mydata[i-1][8],mydata[i-1][9],mydata[i-1][10], mydata[i-1][11], mydata[i-1][12]]])

     }
   

  //----------------

  // var type="csv";

  // const mimeTypes = { csv: MimeType.MICROSOFT_EXCEL, pdf: MimeType.PDF };

  
  // Logger.log("Number of sheets is given by")
  // Logger.log(ss.getSheets())



  // Logger.log("ss id and sheet id is given by")
  // Logger.log(ss.getId());
  // Logger.log(sheet.getSheetId())



  // let url2 = null;
  
  // if (type == "csv") {
  //   url2 = `https://docs.google.com/spreadsheets/d/${ss.getId()}/gviz/tq?tqx=out:csv&gid=${sheet.getSheetId()}`;
  //   Logger.log(url2)

  // } else if (type == "pdf") {
    
  //   url2 = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sheet.getSheetId()}`;
  // }

  // if (url2) {
  //   Logger.log("inside url2 if")
  //   const blob = UrlFetchApp.fetch(url2, {
  //           headers: { authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  //           muteHttpExceptions:true,
  //         }).getBlob();

  //   // var responseCode=blob.getResponseCode()
  //   // var responseBody= blob.getContentText();
  //   // Logger.log(responseCode);
  //   // Logger.log(responseBody);

  //   Logger.log(ScriptApp.getOAuthToken());
  //   Logger.log( Utilities.base64Encode(blob.getBytes()));
  //   return {
  //     data:
  //       `data:${mimeTypes[type]};base64,` +
  //       Utilities.base64Encode(blob.getBytes()),
  //     filename: `${sheet.getSheetName()}.${type}`,
  //   };
  // }
  // return { data: null, filename: null };

 return getFileUrl();
}

// ---------------------------======================================
// function downloadSpreadshee(){
//     var outputDocument = DocumentApp.create('My custom csv file name'); 
//     var content = getCsv();
//     var textContent = ContentService.createTextOutput(content);
//     textContent.setMimeType(ContentService.MimeType.CSV);
//     textContent.downloadAsFile("4NocniMaraton.csv");
//     return textContent;


// }

//======================================


function createDataUrl() {

//----------------------------------------
  var type="csv";

  const mimeTypes = { csv: MimeType.MICROSOFT_EXCEL, pdf: MimeType.PDF };

  const id= "1FfpxMnyfdrzXDbd21KSfvKZT_4j7CJj3xvfBCW3cWfI"
  const ss= SpreadsheetApp.openById(id);
  //var url="https://docs.google.com/spreadsheets/d/1FfpxMnyfdrzXDbd21KSfvKZT_4j7CJj3xvfBCW3cWfI/edit#gid=1060124669"



 // const ss = SpreadsheetApp.openByUrl(url);
  Logger.log(ss.getSheets())

  //const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report");
  const sheet = ss.getSheetByName("Report");
  Logger.log(sheet);

  Logger.log("ss id and sheet id is given by")
  Logger.log(ss.getId());
  Logger.log(sheet.getSheetId())

  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = ss.getActiveSheet();

  // Logger.log("ss id2 and sheet2 id is given by")
  // Logger.log(ss.getId());
  // Logger.log(sheet.getSheetId())


  let url2 = null;
  
  if (type == "csv") {
    url2 = `https://docs.google.com/spreadsheets/d/${ss.getId()}/gviz/tq?tqx=out:csv&gid=${sheet.getSheetId()}`;
    Logger.log(url2)

  } else if (type == "pdf") {
    
    url2 = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sheet.getSheetId()}`;
  }

  if (url2) {
    const blob = UrlFetchApp.fetch(url2, {
      headers: { authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    }).getBlob();
    return {
      data:
        `data:${mimeTypes[type]};base64,` +
        Utilities.base64Encode(blob.getBytes()),
      filename: `${sheet.getSheetName()}.${type}`,
       
    };
  }
  return { data: null, filename: null };

  // DriveApp.getFiles() // This is used for automatically detecting a scope of "https://www.googleapis.com/auth/drive.readonly".
}





//===================================================
function GetReport(searchObject){

  Logger.log("SearchObject")
   Logger.log(searchObject)

      




// make new sheet
return makeNewSheet(searchObject);


}




//==================================================
/**
 * Downloading Data as a CSV File into a folder on Google Drive
 * 
 * 
 * 
 */


function getFileUrl(){

        var filename="Report.csv"

        var folder="1WTgFRF7lUkTYoVJa62VtavqsNA2PC7tM"

        var csv=""
        const id= "1FfpxMnyfdrzXDbd21KSfvKZT_4j7CJj3xvfBCW3cWfI"
        const ss= SpreadsheetApp.openById(id);

        //var value = ss.getSheetByName("Report").getDataRange().getValues();

        var dataSheet = ss.getSheetByName("Report");
        
        // will create an array instead

        var getLastRow = dataSheet.getLastRow();
        Logger.log(getLastRow)
        var return_Array=[];

        for(i = 1; i <= getLastRow; i++)
                {
                  if(dataSheet.getRange(i, 1).getValue() != '')  //(row, col)(2,1)
                  {  
                    var timezone = Session.getScriptTimeZone()
                    var format = 'hh:mm a' // Set date format for output
                    
                    // normalizing the office start time-----

                    if (dataSheet.getRange(i, 10).getValue() !=""){

                        var newDate=dataSheet.getRange(i, 10).getValue();
                    }
                 

                    // normalizing the office end time ------
                   if (dataSheet.getRange(i, 11).getValue() !="") {

                          var newDate2=dataSheet.getRange(i, 11).getValue()


                   }

                    


                    //==================================
                            return_Array.push([
                            dataSheet.getRange(i, 1).getValue(), 
                            dataSheet.getRange(i, 2).getValue(),
                            dataSheet.getRange(i, 3).getValue(),
                            dataSheet.getRange(i, 4).getValue(),
                            dataSheet.getRange(i, 5).getValue(),
                            dataSheet.getRange(i, 6).getValue(),  
                            Utilities.formatDate(new Date(dataSheet.getRange(i, 7).getValue()),"IST", "yyyy-MM-dd"),
                            dataSheet.getRange(i, 8).getValue(),
                            dataSheet.getRange(i, 9).getValue(),

                            dataSheet.getRange(i, 10).getValue() !=""? Utilities.formatDate(newDate, timezone, format):"N/A",
                            dataSheet.getRange(i, 11).getValue() !=""? Utilities.formatDate(newDate2,timezone, format):"N/A"

                                                                 
                    ]);
                  }
                }  
        // return return_Array;  
   Logger.log("return Array")
   Logger.log(return_Array);
  //return return_Array;





        //=====================================


        return_Array.forEach(function(e){

          csv+= e.join(",")+ "\n"
        })

        var url =DriveApp.getFolderById(folder).createFile(filename, csv, MimeType.CSV)
                .getDownloadUrl().replace("?e=download&gd=true", "")


return url;

}





//===================================================

function getStatusOccurance(){

var ss= SpreadsheetApp.openByUrl(url);
var ws= ss.getSheetByName("Data");
var list= ws.getRange(2,9, ws.getRange("I1").getDataRegion().getLastRow(),1).getValues();


var dataObject={};
  
  for (i=0;i<list.length-1;i++){ list[i]=list[i].toString()}
  
 
  const distinct=[... new Set(list)]
  
  
  distinct.forEach(function(read){
    
  
    var count=0;

    

   for (var i=0;i<list.length-1;i++){
    
      if (read == list[i] && read !=[]) { count=count+1;}
    
   }
    
    if (read != "") {
      
      dataObject[read]=count;
      Logger.log(dataObject)
    }
  
  
  });
  
  return (dataObject);
  




}
//=========================================


function getOptStatus(){

var ss= SpreadsheetApp.openByUrl(url);
var ws= ss.getSheetByName("Data");
var list= ws.getRange(2,10, ws.getRange("J1").getDataRegion().getLastRow(),1).getValues();


var dataObject={};
  
  for (i=0;i<list.length-1;i++){ list[i]=list[i].toString()}
  
 
  const distinct=[... new Set(list)]
  
  
  distinct.forEach(function(read){
    
  
    var count=0;

    

   for (var i=0;i<list.length-1;i++){
    
      if (read == list[i] && read !=[]) { count=count+1;}
    
   }
    
    if (read != "") {
      
      dataObject[read]=count;
      Logger.log(dataObject)
    }
  
  
  });
  
  return (dataObject);
  




}


//========================================

function LocationStatusCount(){

var ss= SpreadsheetApp.openByUrl(url);
var ws= ss.getSheetByName("Data");


var list= ws.getRange(2,7, ws.getRange("G1").getDataRegion().getLastRow(),1).getValues();


var dataObject={};
  
  for (i=0;i<list.length-1;i++){ list[i]=list[i].toString()}
  
 
  const distinct=[... new Set(list)]
  
  
  distinct.forEach(function(read){
    
  
    var count=0;

    

   for (var i=0;i<list.length-1;i++){
    
      if (read == list[i] && read !=[]) { count=count+1;}
    
   }
    
    if (read != "") {
      
      dataObject[read]=count;
      Logger.log(dataObject)
    }
  
  
  });
  
  return (dataObject);
  




}


// =======================================  Get Data for time range..










// BackEnd Functions end here
