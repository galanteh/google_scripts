var presentation_template_id = 'YOUR_ID'
var ss_id = 'YOUR_ID'
var parent_folder_id = 'YOUR_ID'
var folder_name = 'Certificados'

// https://stackoverflow.com/questions/1988349/array-push-if-does-not-exist
// check if an element exists in array using a comparer function
// comparer : function(currentElement)
Array.prototype.inArray = function(comparer) { 
    for(var i=0; i < this.length; i++) { 
        if(comparer(this[i])) return true; 
    }
    return false; 
}; 

// adds an element to the array if it does not already exist using a comparer 
// function
Array.prototype.pushIfNotExist = function(element, comparer) { 
    if (!this.inArray(comparer)) {
        this.push(element);
    }
};

// Get the name of the certificates
function get_names(verbose)
{
  if (verbose == null) { verbose = false }
  var ss = SpreadsheetApp.openById(ss_id);
  var sheet = ss.getSheetByName("Names");
  var num1 = 0
  var row = 1
  var column = 1
  var cell_value = ""
  var list_values = []
  do {
    num1++;
    cell_value = sheet.getRange(num1, column).getValue();
    if (verbose) { 
      Logger.log('Row Value: ' +  cell_value);
    }
    if (cell_value.trim() != "") { list_values.push(cell_value) };
  } while (cell_value.trim() != ""); 
    if (verbose) { 
      Logger.log('Values: ' + list_values);
    }
  return list_values
}


function duplicate_template(presentationId, name, folder_cert) {
  var file_pptx = DriveApp.getFileById(presentationId);
  var new_file_pptx = file_pptx.makeCopy(name, folder_cert)
  return new_file_pptx;  
}

// Save all the files in the folder to a zip file as PDF
function zip_folder(source_folder)
{
  var folderId = source_folder.getId();
  var files = source_folder.getFiles();
  var blobs = [];
  var fileBlob = '';
  var file = '';
  Logger.log('Coverting certificates to PDF ...');
  while(files.hasNext()){
    file = files.next();
    fileBlob = file.getAs("application/pdf");    
    blobs.pushIfNotExist(fileBlob, function(e) { 
      return e.getName() === fileBlob.getName(); 
    });
  }
  Logger.log('Compressing all PDF files into one Zip file ...');
  var zippedFolder = Utilities.zip(blobs, folder_name + '_pdfs.zip');
  source_folder.createFile(zippedFolder);

}

// MAIN
function main(){
  var names = get_names()
  var parent_folder = DriveApp.getFolderById(parent_folder_id);
  var folder_cert = parent_folder.createFolder(folder_name);
  Logger.log('Rows to process:' + names.length)
  for (var num = 0; num < names.length + 1; num++) {  
    var name = names[num]
    Logger.log('Processing ' + num + '/' + names.length + ' Certificate to: ' + name)
    var new_file = duplicate_template(presentation_template_id, name, folder_cert)
    Logger.log('ID New File:' + new_file.getId())
    var slide = SlidesApp.openById(new_file.getId())
    slide.replaceAllText("#NOMBRE#", name);
    slide.saveAndClose();
  }
  zip_folder(folder_cert)
}

