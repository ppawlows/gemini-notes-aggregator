const SOURCE_FOLDER_NAME = 'Meet Recordings';
const FILE_PREFIX = 'TRT';
const FILE_SUFFIX = 'Notes by Gemini';

/**
 * Appends cleaned Gemini-generated notes to the beginning of the current document.
 * This version inserts formatted content by copying document elements directly,
 * and sorts source files by creation date (oldest first) so newest appear at top of target.
 */
function appendNewTRTNotes() {
  const folder = getFolderByName(SOURCE_FOLDER_NAME);
  const fileIterator = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  const targetDoc = DocumentApp.getActiveDocument();
  const body = targetDoc.getBody();
  const docProperties = PropertiesService.getDocumentProperties(); // Using DocumentProperties as script is bound

  let filesToProcess = [];

  // Collect all relevant files into an array for sorting
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    const fileName = file.getName();
    const fileId = file.getId();

    // Apply filters before adding to array
    if (!fileName.startsWith(FILE_PREFIX) || !fileName.endsWith(FILE_SUFFIX)) continue;
    if (fileId === targetDoc.getId()) continue;
    if (docProperties.getProperty(fileId) === 'processed') continue; // Uncommented to prevent reprocessing

    filesToProcess.push(file);
  }

  // Sort files by creation date in ascending order (oldest first)
  // This ensures that when prepended (inserted at index 0), the newest will end up at the very top.
  filesToProcess.sort((a, b) => a.getDateCreated().getTime() - b.getDateCreated().getTime());

  let filesProcessedCount = 0;

  for (const file of filesToProcess) {
    const fileName = file.getName();
    const fileId = file.getId();

    try {
      Logger.log(`Processing file for formatted copy: ${fileName} (ID: ${fileId})`);

      const cleanTitle = fileName
        .replace(FILE_SUFFIX, '')         // Remove suffix
        .replace(/\s*-\s*$/, '')          // Remove trailing dash and spaces
        .trim();

      const headerText = `📄 ${cleanTitle}`;

      // Insert header for the new notes at the beginning (index 0)
      // Elements are inserted at index 0, pushing existing content down.
      // So, insert header first, then call copyFormattedContent to insert content at index 1.
      const headerParagraph = body.insertParagraph(0, headerText).setHeading(DocumentApp.ParagraphHeading.HEADING2);

      // Call the function to copy formatted content directly into the target document
      // The content will be inserted right after the header, at index 1.
      copyFormattedContent(fileId, body, 1);

      // Add a blank line *after* the inserted content for spacing between entries
      body.insertParagraph(0, '\n');

      docProperties.setProperty(fileId, 'processed');
      filesProcessedCount++;
      Logger.log(`Successfully appended formatted notes from: ${fileName}`);

    } catch (e) {
      // Log full error stack for better debugging
      Logger.log(`Error processing file ${fileName} (ID: ${fileId}): ${e.message} \nStack: ${e.stack}`);
    }
  }

  targetDoc.saveAndClose();
  Logger.log(`Script finished. ${filesProcessedCount} new files were processed and prepended.`);
}


/**
 * Copies formatted content (paragraphs, lists, tables, images) from a source Google Doc
 * to a specific index in a target document's body.
 * This uses a workaround of exporting to DOCX and re-importing to handle problematic
 * Gemini-generated sections.
 *
 * @param {string} fileId The ID of the source Google Doc file.
 * @param {GoogleAppsScript.Document.Body} targetBody The body object of the target document.
 * @param {number} insertIndex The index at which to insert the content in the targetBody.
 */
function copyFormattedContent(fileId, targetBody, insertIndex) {
  const token = ScriptApp.getOAuthToken();
  const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document`;

  let response;
  try {
    response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true // Allow checking response code for errors
    });
  } catch (e) {
    Logger.log(`Error during export fetch for file ${fileId}: ${e.message}`);
    return; // Exit function on fetch error
  }

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    Logger.log(`Failed to export file ${fileId}. Status: ${responseCode}, Message: ${response.getContentText()}`);
    return; // Exit function if export failed
  }

  const blob = response.getBlob();
  let tempFileId = null; // Store the ID of the temporary file for proper cleanup

  try {
    // Requires Drive Advanced Service (v3) to be enabled.
    // Ensure 'Drive' service is set to v3 in Apps Script project settings.
    const tempFile = Drive.Files.create( // CORRECTED: Changed .create to .insert
      {
        name: `temp_doc_from_gemini_${Date.now()}`, // Unique title
        mimeType: MimeType.GOOGLE_DOCS,
        parents: [{ id: DriveApp.getRootFolder().getId() }] // Optional: specify a folder
      },
      blob
    );
    tempFileId = tempFile.id; // Save ID for cleanup

    const tempDoc = DocumentApp.openById(tempFileId);
    const tempBody = tempDoc.getBody();
    const numElements = tempBody.getNumChildren();

    if (numElements === 0) {
      Logger.log(`Temporary document ${tempFileId} is empty. Skipping content copy.`);
      return;
    }

    // Iterate from last to first to insert elements at a fixed index correctly
    for (let i = numElements - 1; i >= 0; i--) {
      const element = tempBody.getChild(i);
      const elementType = element.getType();

      switch (elementType) {
        case DocumentApp.ElementType.PARAGRAPH:
          targetBody.insertParagraph(insertIndex, element.asParagraph().copy());
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          targetBody.insertListItem(insertIndex, element.asListItem().copy());
          break;
        case DocumentApp.ElementType.TABLE:
          targetBody.insertTable(insertIndex, element.asTable().copy());
          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          // InlineImage is usually a child of a Paragraph.
          // This part can be more complex if images are not simple inline.
          try {
            const paragraphWithImage = element.asParagraph();
            for (let j = 0; j < paragraphWithImage.getNumChildren(); j++) {
              const child = paragraphWithImage.getChild(j);
              if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
                targetBody.insertImage(insertIndex, child.asInlineImage().copy());
                break; // Assuming one image per paragraph for simplicity
              }
            }
          } catch (e) {
            Logger.log(`Could not copy inline image from paragraph: ${e.message}`);
          }
          break;
        case DocumentApp.ElementType.HORIZONTAL_RULE:
          targetBody.insertHorizontalRule(insertIndex);
          break;
        case DocumentApp.ElementType.PAGE_BREAK:
          targetBody.insertPageBreak(insertIndex);
          break;
        // Add more cases for other element types if needed
        case DocumentApp.ElementType.UNSUPPORTED:
          // For UNSUPPORTED elements, we can't get text or copy directly.
          // Log it and skip, or try to get outer HTML if that makes sense (more complex).
          Logger.log(`Skipping unsupported element type: ${elementType}. Cannot copy directly.`);
          break;
        default:
          // For other unhandled types, try to get text content if available
          Logger.log(`Encountered unhandled element type: ${elementType}. Attempting to copy as plain text if possible.`);
          // Check if getText() method exists on the element before calling it
          if (typeof element.getText === 'function') {
            const textContent = element.getText();
            if (textContent) {
              targetBody.insertParagraph(insertIndex, textContent);
            }
          } else {
            Logger.log(`Element of type ${elementType} does not support getText() method. Skipping.`);
          }
          break;
      }
    }

  } catch (e) {
    Logger.log(`Error during import/read/copy process for file ${fileId}: ${e.message}`);
  } finally {
    // Ensure temporary file is deleted, even if errors occur during try block.
    if (tempFileId) {
      try {
        Drive.Files.remove(tempFileId);
        Logger.log(`Temporary file ${tempFileId} deleted successfully.`);
      } catch (e) {
        Logger.log(`Failed to delete temporary file ${tempFileId}: ${e.message}. Manual cleanup may be required.`);
      }
    }
  }
}

/**
 * Retrieves a Google Drive folder by its name.
 * Throws an error if the folder is not found.
 * @param {string} name The name of the folder to retrieve.
 * @returns {GoogleAppsScript.Drive.Folder} The found folder.
 */
function getFolderByName(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (!folders.hasNext()) {
    throw new Error(`Folder '${name}' not found. Please ensure the folder name is correct.`);
  }
  return folders.next();
}
