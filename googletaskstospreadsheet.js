function getTaskList(list) {
  var taskList  = Tasks.Tasklists.list().getItems();
  for( var i = 0; i < taskList.length; i++) {
    if(taskList[i].getTitle() == list) {
      return taskList[i].getId();
    }
  }
}

function getTasks(tasklist) {
  var list = Tasks.Tasks.list(tasklist);
  
  return list.getItems();
}

function setTaskToSpread() {
  var mod = 2;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  
  
  var list = ss.getRange("A1").getValue();//"Insight";

  if( ss.getRange("A1").isBlank() || ss.getRange("A1").getValue() == "ENTER TEXT HERE") {
    sheet.clear();
    
    ss.getRange("A1").setValue("ENTER TEXT HERE");
    return;
  }
  
  sheet.clear();
  ss.getRange("A1").setValue(list);//"Insight";

  
  var tasklist = getTaskList(list);
  
  var tasks = getTasks(tasklist);
  
  for( var i = 0; i < tasks.length ; i++) {
    var row = i+mod;
    sheet.setActiveCell("A"+row).setValue(tasks[i].getTitle());
    if(tasks[i].getDue() != null)
      sheet.setActiveCell("B"+row).setValue(tasks[i].getDue().substring(0,10));
    if(tasks[i].getCompleted() != null)
      sheet.setActiveCell("C"+row).setValue(tasks[i].getCompleted().substring(0,10));
    
    sheet.setActiveCell("D"+row).setFormula("=IF(C"+row+">B"+row+",C"+row+"-B"+row+",0)");
  }
}
