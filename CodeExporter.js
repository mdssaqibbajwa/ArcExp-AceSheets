/**
 * Fetches project content using UrlFetchApp and the REST API, then writes it to a Google Doc.
 */
function codeExporter(scriptId, fileName, destinationFolder) {
  Logger.log(`Fetching content for project ID: ${scriptId} using UrlFetchApp.`);

  try {
    const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true // Prevents script from stopping on HTTP errors
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode !== 200) {
      Logger.log(`ERROR: Failed to fetch project content. Response Code: ${responseCode}. Response Body: ${responseBody}`);
      return;
    }

    const content = JSON.parse(responseBody);

    if (!content.files || content.files.length === 0) {
      Logger.log(`No files found in project "${fileName}". Skipping doc creation.`);
      return;
    }

    Logger.log(`Creating Google Doc for project: "${fileName}"`);
    //const docName = `Apps Script Code - ${fileName}`;
    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();

    // New styles with improved design
    var headerStyle = {};
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.ITALIC] = true;
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1D8A3E';  // brand green
    headerStyle[DocumentApp.Attribute.SPACING_BEFORE] = 18;  // breathing room before each file block
    headerStyle[DocumentApp.Attribute.SPACING_AFTER] = 4;

    var codeStyle = {};
    codeStyle[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.COURIER_NEW;
    codeStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
    codeStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#2D2D2D';  // softer than pure black
    codeStyle[DocumentApp.Attribute.BOLD] = false;
    codeStyle[DocumentApp.Attribute.ITALIC] = false;
    codeStyle[DocumentApp.Attribute.SPACING_BEFORE] = 2;
    codeStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;

    // [DOWNLOAD CODE AS TEXT] Accumulates plain-text content for client-side .txt download
    var textLines = [];
    content.files.forEach(file => {
      var ext = file.type === 'SERVER_JS' ? 'gs' : 'json';

      body.appendParagraph(`// FILE: ${file.name}.${file.type === 'SERVER_JS' ? 'gs' : 'json'}`)
        .setAttributes(headerStyle);

      const codeBlock = body.appendParagraph(file.source);
      codeBlock.setAttributes(codeStyle);
      body.appendParagraph('');

      // [DOWNLOAD CODE AS TEXT] Mirror each file into the plain-text accumulator
      textLines.push('// FILE: ' + file.name + '.' + ext);
      textLines.push(file.source);
      textLines.push('');
    });

    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(destinationFolder);

    Logger.log(`Successfully created and moved document: ${docFile.getName()} (ID: ${docFile.getId()})`);

    return {
      success: true,
      docUrl: docFile.getUrl(),
      docName: docFile.getName(),
      appsScriptContent: textLines.join('\n')  // [DOWNLOAD CODE AS TEXT]
    };

  } catch (e) {
    Logger.log(`CRITICAL ERROR exporting file "${fileName}" (ID: ${scriptId}). Error: ${e.toString()}`);
    Logger.log(`Stack Trace: ${e.stack}`);
    return { success: false, error: e.message };
  }
}