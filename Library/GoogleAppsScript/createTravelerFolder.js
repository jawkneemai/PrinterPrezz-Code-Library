function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Custom Functions')
    .addItem('Create Traveler Folder', 'createTravelerFolder')
    .addToUi();
}

function createTravelerFolder() {
  var projectsFolder = DriveApp.getFolderById('1lNxvM7WIKcGfM-XR-nAbP6nT0NtYPC7T')
  var travelerFolders = projectsFolder.getFolders();
  var templateFolder = DriveApp.getFolderById('1_coLyOR-yEEWibsedQZsiVZEmcTWO0Y3');
  var travTemplateFolder = DriveApp.getFolderById('1_h_qUM6mej0K2nMZ0r-mvTrxhZXUsBrn');
  var labelTemplateFolder = DriveApp.getFolderById('1alLHPJnbJkPQ9jBmqF9AaGTudW5DHC-s');
  
  var folder = travelerFolders.next();
  var newestFolder = folder;
  
  // Find latest traveler folder
  /* This algorithm finds most recently created traveler folder and determines next traveler number from that
  while (travelerFolders.hasNext()) {
    var folder = travelerFolders.next();
    var tempDate = folder.getDateCreated();
    if (tempDate.getTime() > newestFolder.getDateCreated().getTime()) {
      newestFolder = folder;
    } 
  }


  // Create new traveler folder at latest traveler folder + 1
  var travNum = parseInt(newestFolder.getName().split(' ')[0].substring(1));
  travNum += 1;
  var folderName = 'T' + travNum.toString();

  */

  var folderName = 'TXXXX';
  var newestFolder = projectsFolder.createFolder(folderName);



  // Get all files 
  var templateFiles = templateFolder.getFolders();
  var travTemplateFiles = travTemplateFolder.getFiles();
  var labelTemplateFiles = labelTemplateFolder.getFiles();

  // Create copies of folders in new traveler folder
  while (templateFiles.hasNext()) {
    var tempFolder = templateFiles.next();
    newestFolder.createFolder(tempFolder.getName());
  }

  // Create copies of traveler files in traveler folder, label files in label folder, folder path in Build Files folder.
  var newTravFolder = '';
  var newLabelFolder = '';
  var newBuildFilesFolder = '';
  var newFolderFolders = newestFolder.getFolders();
  while (newFolderFolders.hasNext()) {
    var tempFolder = newFolderFolders.next();
    if (tempFolder.getName() == 'Traveler Forms') {
      newTravFolder = tempFolder;
    } else if (tempFolder.getName() == 'Labels') {
      newLabelFolder = tempFolder;
    } else if (tempFolder.getName() == 'Build Files') {
      newBuildFilesFolder = tempFolder;
      var rfqFolder = newBuildFilesFolder.createFolder('RFQXXXX-X');
      rfqFolder.createFolder('Part Files');
    }
  }

  while (travTemplateFiles.hasNext()) {
    travTemplateFiles.next().makeCopy(newTravFolder);
  }

  while (labelTemplateFiles.hasNext()) {
    labelTemplateFiles.next().makeCopy(newLabelFolder);
  }
  




  return;
}