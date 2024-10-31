function syncGoogleTaskToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignment Dashboard'); // Update sheet name
  var taskLists = Tasks.Tasklists.list(); // Get all task lists
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastTaskAddedTime = scriptProperties.getProperty('LAST_TASK_TIME') || 0; // Get the last task time or set to 0

  // Log task lists checked
  Logger.log('Checking task lists...');
  var allTasksInSheet = {};
  var existingTasksInGoogle = {};

  // Track existing tasks in the sheet
  var assignments = sheet.getRange('C16:C').getValues(); // Get tasks from the sheet
  for (var row = 0; row < assignments.length; row++) {
    var assignmentInSheet = assignments[row][0];
    if (assignmentInSheet) {
      allTasksInSheet[assignmentInSheet] = true; // Mark task as existing in the sheet
    }
  }

  var totalTasksInGoogle = 0;

  for (var i = 0; i < taskLists.items.length; i++) {
    var taskList = taskLists.items[i];
    var tasks = Tasks.Tasks.list(taskList.id).items; // Get tasks from each task list
    Logger.log('Task List: ' + taskList.title + ' has ' + (tasks ? tasks.length : 0) + ' tasks.');

    if (tasks && tasks.length > 0) {
      totalTasksInGoogle += tasks.length; // Count total tasks in Google Tasks
      for (var j = 0; j < tasks.length; j++) {
        var task = tasks[j];
        existingTasksInGoogle[task.title] = { 
          listName: taskList.title,  // Store the list name
          details: task.notes || '', // Store task details (notes)
          dueDate: task.due ? new Date(task.due).toLocaleDateString("en-GB") : '' // Store due date
        }; // Mark task as existing in Google Tasks
      }
    }
  }

  // Only proceed if there is a change in the number of tasks
  if (totalTasksInGoogle !== Object.keys(allTasksInSheet).length) {
    // Check for missing tasks in the sheet
    for (var taskTitle in allTasksInSheet) {
      if (allTasksInSheet.hasOwnProperty(taskTitle) && !existingTasksInGoogle[taskTitle]) {
        Logger.log('Marking task as completed in sheet: ' + taskTitle);
        var rowToUpdate = findRowByTaskTitle(sheet, taskTitle);
        if (rowToUpdate) {
          sheet.getRange(rowToUpdate, 5).setValue('Completed'); // Update status in Column E
        }
      }
    }

    // Now add any new tasks
    for (var taskTitle in existingTasksInGoogle) {
      if (existingTasksInGoogle.hasOwnProperty(taskTitle) && !allTasksInSheet[taskTitle]) {
        Logger.log('Adding new task: ' + taskTitle);
        var taskData = existingTasksInGoogle[taskTitle]; // Get the task's details
        
        // Insert a new row at row 16
        sheet.insertRowBefore(16); // Inserts a row before row 16

        // Insert the task information into respective columns
        sheet.getRange(16, 2).setValue(taskData.listName); // Column B: Class (task list name)
        sheet.getRange(16, 3).setValue(taskTitle); // Column C: Assignment
        sheet.getRange(16, 4).setValue(taskData.details); // Column D: Priority (task details/notes)
        sheet.getRange(16, 6).setValue(taskData.dueDate); // Column F: Due Date

        // Set the default status to "Not Started" in Column E
        sheet.getRange(16, 5).setValue("Not Started"); // Column E: Status

        // Merge columns G and H for Days Left
        sheet.getRange("G16:H16").merge();
        // Merge columns I and J for any other purpose you may have
        sheet.getRange("I16:J16").merge();

        // Set the formula for Days Left in merged cells G16 and H16
        sheet.getRange("G16").setFormula('=IF(ISBLANK(F16), "", IF(F16 < TODAY(), "Overdue", DATEDIF(TODAY(), F16, "D")))');

        // Clear row 16, columns K to O without interfering with merged cells
        sheet.getRange("K16:O16").clearContent();

        // Preserve values in columns K to O from row 16 by moving the contents down
        sheet.getRange("K17:O17").setValues(sheet.getRange("K16:O16").getValues());

        // Update the last task time to the current task's time
        var taskCreationTime = new Date().getTime(); // You can modify this as per your logic
        scriptProperties.setProperty('LAST_TASK_TIME', taskCreationTime);
      }
    }
  } else {
    Logger.log('No changes in the number of tasks; no action taken.');
  }

  // Log the completion of the sync process
  Logger.log('Sync process completed.');
}

// Helper function to find the row number of a task title in the sheet
function findRowByTaskTitle(sheet, taskTitle) {
  var assignments = sheet.getRange('C16:C').getValues();
  for (var row = 0; row < assignments.length; row++) {
    if (assignments[row][0] === taskTitle) {
      return row + 16; // Return the actual row number in the sheet
    }
  }
  return null; // Task not found
}