// -------------------------------
// HISTORY
// https://medium.com/@ocangelo/automatically-backup-kindle-scribe-notebooks-to-google-drive-4937b5b30a2f
//
// 1.2
//  FIXED:
//    - parsing new format of emails (November 2024) from amazon to grab notebooks, pages and text files from "convert to text"
//  IMPORTANT:  
//    - as of November 2024 PAGES can no longer be recognized because of the new email format amazon is using,
//      that means pages and notebooks will be named the same and override each other is DELETE_OLD_BACKUPS is true
//
// 1.1.2
//    NEW: 
//      - added EMAIL_TRASH setting
// 1.1
//    FIXED: 
//      - special characters in notebooks names were breaking the script
//    NEW:  
//      - added 3 new settings related to all processed emails:
//          + apply a label to all of them
//          + mark them as read
//          + archive them
// -------------------------------

// -------------------------------
// SETTINGS
// -------------------------------

// where to backup
const  DRIVE_FOLDER = "Documents/Kindle Scribe Notes Backup"

// only keep the most recent backup of each notebook if true (it will delete older versions of files with the same name) - pages will never be deleted
// IMPORTANT: as of November 2024 PAGES can no longer be recognized because of the new email format amazon is using,
//            that means pages and notebooks will be named the same and override each other is DELETE_OLD_BACKUPS is true
const  DELETE_OLD_BACKUPS = true

// emails newer than this interval will be scanned
const  EMAIL_NEWER_THAN = "1h"

// apply this label to the email after backup
const  EMAIL_APPLY_LABEL = "_kindleBackup"

// mark email as read after backup
const  EMAIL_MARK_AS_READ = true

// archive email after backup
const  EMAIL_ARCHIVE = false

// trash email after backup
const  EMAIL_TRASH = true


// -------------------------------
// INTERNAL SETTINGS
// -------------------------------

// version reference
const  VERSION = "1.2"

// which emails to consider
const  EMAIL_FILTER = `from:do-not-reply@amazon.com "from your Kindle" newer_than:` + EMAIL_NEWER_THAN

// emails matching this subject are considered notebooks
const  NOTEBOOK_FILTER = "you sent a file \"(.+)\" from your kindle"

// enable temporarily for additional logging
const  PRINT_DETAILED_LOG = false


// -------------------------------
// DEPRECATED SETTINGS
// -------------------------------
// as of November 2024 amazon changed the format of emails and pages and notebooks are sent with the same format
// so we can't separate them anymore and these settings won't work

// add an additional suffix to the file name for pages
const  PAGE_SUFFIX = " PAGE"

// emails matching this subject are considered pages
const  PAGE_FILTER = "you sent some pages of \"(.+)\" from your kindle"

// -------------------------------
// CODE
// -------------------------------

function PrintLog(msg, detailed = false)
{
  if ((!detailed) || PRINT_DETAILED_LOG)
  {
    Logger.log(msg);
  }
}

// structure representing notes we read from emails
// each map uses the notebook/page name as a key and a list of urls as a value
function Notes() {
  this.notebooks = new Map();
  this.pages = new Map();
  this.emailThreads = [];

  this.hasNotes = function() { return this.notebooks.size > 0 || this.pages.size > 0 }
}


function main() {
  PrintLog("--- Running version " + VERSION)

  notes = new Notes();

  PrintLog("--- READING EMAILS")
  getNotesFromEmails(notes);

  if (notes.hasNotes())
  {
    PrintLog("--- FOUND EMAILS")
    const folder = createFolder(DRIVE_FOLDER);
 
    PrintLog("--- UPLOADING NOTEBOOKS")
    uploadNotesToFolder(notes.notebooks, folder, DELETE_OLD_BACKUPS);
    
    PrintLog("--- UPLOADING PAGES")
    uploadNotesToFolder(notes.pages, folder, false);

    PrintLog("--- MARKING EMAILS")
    markEmails(notes.emailThreads)
  }

  PrintLog("--- DONE")
}

// retrieve all notes from gmail
function getNotesFromEmails(notes) {
  GmailApp.search(EMAIL_FILTER).forEach((thread) => {
    thread.getMessages().forEach((message) => {
      if (!message.isInTrash())
        if (getNotesFromMessage(message, notes))
          notes.emailThreads.push(thread);
    })
  })
}


// downloads and uploads all notes to the given folder applying prefix to the name
function uploadNotesToFolder(notes, folder, deleteOldBackups = false) {
  notes.forEach((noteUrls, noteName) => {
    noteUrls.forEach((url) => {
      const name = noteName + "." + getNoteExtFromUrl(url)
      PrintLog("processing \"" + name + "\" from " + url)

      // remove old backup
      if (deleteOldBackups)
      {
        var oldBackups = folder.getFilesByName(name);
        while (oldBackups.hasNext()) {
          PrintLog("deleting old backup for " + name)
          var thisFile = oldBackups.next();
          thisFile.setTrashed(true);
        };
      }

      // fetch the file
      const options = {}//{muteHttpExceptions: true};
      const noteBlob = UrlFetchApp.fetch(url, options)
      if (PRINT_DETAILED_LOG)
        PrintLog(noteBlob.getContentText());

      // create the new file in gdrive
      let noteFile = folder.createFile(noteBlob.getBlob())

      noteFile.setName(name)
      noteFile.setDescription("Backup of " + name + " from Kindle Scribe")
      PrintLog("Backup of " + name + " from Kindle Scribe COMPLETED")
    })
  })
}

// apply labels, mark as read, archive etc.
function markEmails(emailThreads)
{
  if (EMAIL_APPLY_LABEL != "")
  {
    var label = GmailApp.getUserLabelByName(EMAIL_APPLY_LABEL);
    if (!label)
    {
      PrintLog("creating label " + EMAIL_APPLY_LABEL)
      label = GmailApp.createLabel(EMAIL_APPLY_LABEL);
    }

    if (label)
    {
      PrintLog("applying label " + EMAIL_APPLY_LABEL)
      emailThreads.forEach(email => {
        email.addLabel(label);
      });
    }
    else
      PrintLog("something went wrong with the creation of the label")
  } 

  if (EMAIL_MARK_AS_READ)
  {
    PrintLog("marking emails as read")
    GmailApp.markThreadsRead(emailThreads);
  }

  if (EMAIL_ARCHIVE)
  {
    PrintLog("marking emails as archived")
    GmailApp.moveThreadsToArchive(emailThreads);
  }
  
  if (EMAIL_TRASH)
  {
    PrintLog("trashing emails")
    GmailApp.moveThreadsToTrash(emailThreads);
  }
}

// grabs file extension from url
function getNoteExtFromUrl(url) {
  var pattern = /kindle-content-requests-prod.s3.amazonaws.com\/[\w\d-]+\/[\d\w\s\-%]+\.([\w]+)\?/i
  const matches = url.match(pattern)
  if (matches) {
    return matches[1];
  }

  PrintLog("cannot retrieve note extension from " + url)

  return "unknown"
}

// retrieve the url from the body of the message
function getNoteUrlsFromMessage(message) {
  const urlPattern = /https%3A%2F%2Fkindle-content-requests-prod[\w\-\.\%]*Signature%3D[a-zA-Z0-9]{64}/gi

  // replace raw newlines
  const content = message.replace(/=[\n\v\r]+/gi,"")

  const matches = Array.from(content.matchAll(urlPattern), (m) => m[0])

  if (matches.length > 0) {
    PrintLog("found " + matches.length + " links", true)
    const res = matches.map((url) => decodeURIComponent(url))
    res.forEach((url) => {
      PrintLog("found url: " + url)
    })

    return res
  }

  PrintLog("no url found on message")
  return
}


// parse the message and if it matches pattern add urls to noteMap
function parseMessage(message, notesMap, pattern, canDeleteOldBackups = false, nameSuffix = "") {  
  const subject = message.getSubject()

  const matches = pattern.exec(subject);
  if (matches)
  {    
      const name = matches[1] + nameSuffix;
      PrintLog("found message with notes: " + name);

      var mapItem = notesMap.get(name)
      if (mapItem && canDeleteOldBackups)
      {         
        // we cannot upload from multiple emails with the same notebook name
        PrintLog("skipping older upload of " + name)
        return [name, 0]
      }

      const noteUrls = getNoteUrlsFromMessage(message.getRawContent())
      if (noteUrls && noteUrls.length > 0)
      {
        if (!mapItem)
        {
          notesMap.set(name, noteUrls)
        }
        else
        {
          // upload older versions as well
          noteUrls.forEach((url) => mapItem.push(url))
        }

        return [name, noteUrls.length];
      }

      return [name, 0]
  }

  return
}

// retrieve all notebooks and pages from the given email
function getNotesFromMessage(message, notes) {  
  const subject = message.getSubject()

  const notebookPattern = new RegExp(NOTEBOOK_FILTER, "gi");
  const pagesPattern = new RegExp(PAGE_FILTER, "gi");

  if (notebookFound = parseMessage(message, notes.notebooks, notebookPattern, DELETE_OLD_BACKUPS))
  {
    if (notebookFound[1] > 0)
    {
      PrintLog("added notebook " + notebookFound[0]);
      return true;
    }
        
    return false;
  }
  else if (pageName = parseMessage(message, notes.pages, pagesPattern, false, PAGE_SUFFIX))
  {
    if (notebookFound[1] > 0)
    {
      PrintLog("added page " + notebookFound[0]);
      return true;
    }

    return false
  }
  else
  {
    PrintLog("Skipping invalid email with subject: " + subject, true);
    return false;
  }
}

// create a new folder in google drive
function createFolder(path) {
   // Start at the root.
  let currentFolder = DriveApp.getRootFolder();

  const parts = path.split('/');
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
   
    // Skip empty parts, could happen with leading or trailing slashes.
    if (!part)
      continue;
   
    // Try to find the next part of the path in the current folder.
    let nextFolder = null;
    const folders = currentFolder.getFoldersByName(part);
   
    // Check if the folder exists.
    while (folders.hasNext()) {
      const folder = folders.next();

      // If the folder is found, set it as the next folder to process.
      if (folder.getName() === part) {
        nextFolder = folder;
        break;
      }
    }
   
    // If the folder doesn't exist, create it.
    if (!nextFolder)
      nextFolder = currentFolder.createFolder(part);
   
    // Update the current folder to the new folder (either found or created).
    currentFolder = nextFolder;
  }
 
  return currentFolder;
}
