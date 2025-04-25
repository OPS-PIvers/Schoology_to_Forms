/**
 * IMSCC to Google Form Converter
 * 
 * This script converts IMS Common Cartridge (IMSCC) files to Google Forms
 * and adds form information to the active spreadsheet.
 */

// Global configuration
const CONFIG = {
  TEMP_FOLDER_NAME: 'IMSCC_Temp',
  SPREADSHEET_HEADERS: ['Form Title', 'Form URL', 'Edit URL', 'Number of Questions', 'Creation Date']
};

/**
 * Creates a menu item for the add-on.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('IMSCC Converter')
      .addItem('Convert IMSCC to Google Form', 'showSidebar')
      .addToUi();
}

/**
 * Shows a sidebar with file picker functionality.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('IMSCC to Google Form Converter');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Finds IMSCC files in the user's Drive.
 * 
 * @return {Array} - Array of IMSCC file objects with id and name
 */
function findIMSCCFiles() {
  console.log("Finding IMSCC files in Drive...");
  const files = [];
  
  try {
    // Search for files with .imscc extension, with pagination
    const query = "fileExtension = 'imscc'";
    const fileIterator = DriveApp.searchFiles(query);
    
    // Limit to 50 files for better performance
    let count = 0;
    const maxFiles = 50;
    
    while (fileIterator.hasNext() && count < maxFiles) {
      const file = fileIterator.next();
      files.push({
        id: file.getId(),
        name: file.getName()
      });
      count++;
    }
    
    console.log(`Found ${files.length} IMSCC files`);
    return files;
  } catch (error) {
    console.log(`Error finding IMSCC files: ${error.message}`);
    return [];
  }
}

/**
 * Main function to convert an IMSCC file to Google Forms.
 * 
 * @param {string} fileId - The ID of the IMSCC file in Google Drive
 * @return {Object} - Result information including form URLs
 */
function convertIMSCCToForms(fileId) {
  console.log(`Starting conversion of IMSCC file: ${fileId}`);
  
  // Create temp folder first so we have somewhere to work even if errors occur
  let tempFolder = null;
  
  try {
    // Get file from Drive
    const file = DriveApp.getFileById(fileId);
    if (!file) {
      throw new Error('File not found');
    }
    
    console.log(`Working with file: ${file.getName()} (${file.getSize()} bytes, ${file.getMimeType()})`);
    
    // Check file extension
    if (!file.getName().toLowerCase().endsWith('.imscc')) {
      throw new Error('Not an IMSCC file. Please select a file with .imscc extension.');
    }
    
    // Create temp folder
    tempFolder = createTempFolder();
    console.log(`Created temporary folder: ${tempFolder.getId()}`);
    
    // Extract IMSCC file
    console.log("Extracting IMSCC file...");
    const extractedFiles = extractIMSCC(file, tempFolder);
    console.log(`Extraction complete. Got ${extractedFiles.length} files.`);
    
    // Find and parse manifest
    console.log("Finding and parsing manifest...");
    const manifest = findAndParseManifest(tempFolder);
    
    // Find quizzes in the manifest
    console.log("Finding quizzes in manifest...");
    const quizzes = findQuizzesInManifest(manifest, tempFolder);
    console.log(`Found ${quizzes.length} potential quizzes in manifest`);
    
    // Create Google Forms
    console.log("Creating Google Forms...");
    const formIds = createGoogleForms(quizzes);
    console.log(`Created ${formIds.length} Google Forms`);
    
    // Update spreadsheet with form information
    if (formIds.length > 0) {
      console.log("Updating spreadsheet with form information...");
      updateSpreadsheetWithFormInfo(formIds);
    }
    
    // Return results
    return {
      success: true,
      message: `Successfully converted ${formIds.length} quizzes to Google Forms.`,
      formIds: formIds
    };
    
  } catch (error) {
    console.log(`Error converting IMSCC file: ${error.message}`);
    console.log(`Stack trace: ${error.stack}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  } finally {
    // Clean up temp folder (if it was created) regardless of success/failure
    if (tempFolder) {
      try {
        cleanupTempFolder(tempFolder);
        console.log('Temporary folder cleaned up.');
      } catch (cleanupError) {
        console.log(`Warning: Could not clean up temporary folder: ${cleanupError.message}`);
      }
    }
  }
}

/**
 * Creates a temporary folder for IMSCC extraction.
 * 
 * @return {Folder} - The created temp folder
 */
function createTempFolder() {
  // Check if temp folder already exists and delete it
  const folderIterator = DriveApp.getFoldersByName(CONFIG.TEMP_FOLDER_NAME);
  if (folderIterator.hasNext()) {
    const existingFolder = folderIterator.next();
    existingFolder.setTrashed(true);
  }
  
  // Create new temp folder
  return DriveApp.createFolder(CONFIG.TEMP_FOLDER_NAME);
}

/**
 * Extracts an IMSCC file (ZIP) to the temp folder.
 * 
 * @param {File} file - The IMSCC file
 * @param {Folder} tempFolder - The temporary folder
 * @return {Array} - Array of extracted files
 */
function extractIMSCC(file, tempFolder) {
  try {
    console.log(`Attempting to extract file: ${file.getName()} (${file.getSize()} bytes)`);
    
    // Get the blob with explicit mime type
    let blob = file.getBlob();
    
    // Make sure we have a valid blob with appropriate MIME type
    if (!blob) {
      throw new Error("Failed to get blob from file");
    }
    
    console.log(`Got blob: ${blob.getName()}, size: ${blob.getBytes().length} bytes, content type: ${blob.getContentType()}`);
    
    // Explicitly set to application/zip if needed
    if (blob.getContentType() !== 'application/zip' && 
        blob.getContentType() !== 'application/x-zip-compressed') {
      console.log("Setting content type to application/zip");
      blob = blob.setContentType('application/zip');
    }
    
    // Try to unzip the file
    console.log("Attempting to unzip the blob...");
    let extracted;
    try {
      extracted = Utilities.unzip(blob);
    } catch (unzipError) {
      console.log(`Initial unzip failed: ${unzipError.message}`);
      
      // Try with a different approach - copy the file to ensure it's fully accessible
      console.log("Trying alternative extraction approach...");
      const tempFile = tempFolder.createFile(blob);
      const newBlob = tempFile.getBlob().setContentType('application/zip');
      extracted = Utilities.unzip(newBlob);
      
      // Delete the temporary copy after extraction
      tempFile.setTrashed(true);
    }
    
    if (!extracted || extracted.length === 0) {
      throw new Error("Extraction produced no files, the IMSCC file may be invalid or empty");
    }
    
    console.log(`Successfully unzipped ${extracted.length} files/folders`);
    
    const extractedFiles = [];
    
    for (let i = 0; i < extracted.length; i++) {
      const extractedBlob = extracted[i];
      
      // Skip directories (usually end with a slash)
      if (extractedBlob.getName().endsWith('/')) {
        continue;
      }
      
      // Handle paths with directories
      const fileName = extractedBlob.getName();
      const parts = fileName.split('/');
      
      // Create subdirectories as needed
      let currentFolder = tempFolder;
      if (parts.length > 1) {
        for (let j = 0; j < parts.length - 1; j++) {
          const folderName = parts[j];
          if (!folderName) continue; // Skip empty folder names
          
          // Check if folder exists
          let nextFolder = null;
          const subFolders = currentFolder.getFoldersByName(folderName);
          if (subFolders.hasNext()) {
            nextFolder = subFolders.next();
          } else {
            nextFolder = currentFolder.createFolder(folderName);
          }
          currentFolder = nextFolder;
        }
      }
      
      // Create the file in the appropriate folder
      const extractedFileName = parts[parts.length - 1];
      if (!extractedFileName) continue; // Skip if filename is empty
      
      try {
        const newFile = currentFolder.createFile(extractedBlob);
        extractedFiles.push(newFile);
        
        console.log(`Extracted file: ${fileName}`);
      } catch (fileError) {
        console.log(`Error creating file ${fileName}: ${fileError.message}`);
      }
    }
    
    console.log(`Total extracted files: ${extractedFiles.length}`);
    return extractedFiles;
    
  } catch (error) {
    console.log(`Error in extractIMSCC: ${error.message}`);
    console.log(`Error stack: ${error.stack}`);
    throw new Error(`Failed to extract IMSCC file: ${error.message}`);
  }
}

/**
 * Finds and parses the imsmanifest.xml file.
 * 
 * @param {Folder} tempFolder - The folder containing extracted IMSCC
 * @return {Document} - XML Document of the manifest
 */
function findAndParseManifest(tempFolder) {
  const manifestFiles = tempFolder.getFilesByName('imsmanifest.xml');
  
  if (!manifestFiles.hasNext()) {
    throw new Error('No imsmanifest.xml found in the IMSCC file.');
  }
  
  const manifestFile = manifestFiles.next();
  const content = manifestFile.getBlob().getDataAsString();
  
  // Parse XML
  try {
    const document = XmlService.parse(content);
    return document;
  } catch (e) {
    console.log("Error parsing manifest XML: " + e.message);
    throw new Error("Failed to parse manifest XML: " + e.message);
  }
}

/**
 * Finds quizzes in the manifest.
 * 
 * @param {Document} manifest - XML Document of the manifest
 * @param {Folder} tempFolder - The folder containing extracted IMSCC
 * @return {Array} - Array of quiz information objects
 */
function findQuizzesInManifest(manifest, tempFolder) {
  const quizzes = [];
  const root = manifest.getRootElement();
  
  try {
    // Add necessary namespaces for proper XML parsing
    const imsNS = XmlService.getNamespace('http://www.imsglobal.org/xsd/imscp_v1p1');
    
    // Find resources
    const resources = root.getChild('resources', imsNS);
    if (!resources) {
      console.log('No resources found in manifest.');
      return quizzes;
    }
    
    const resourceElements = resources.getChildren('resource', imsNS);
    for (const resourceElement of resourceElements) {
      const typeAttr = resourceElement.getAttribute('type');
      if (!typeAttr) continue;
      
      const type = typeAttr.getValue();
      console.log(`Found resource type: ${type}`);
      
      // Look for quiz content - expand types to catch more quiz formats
      if (type === 'imsqti_xmlv1p2' || 
          type === 'imsqti_xmlv2p1' || 
          type.includes('assessment') || 
          type.includes('qti') ||
          type.includes('quiz')) {
        
        const identifierAttr = resourceElement.getAttribute('identifier');
        if (!identifierAttr) continue;
        
        const identifier = identifierAttr.getValue();
        const hrefAttr = resourceElement.getAttribute('href');
        const href = hrefAttr ? hrefAttr.getValue() : '';
        
        // Find associated files
        const files = [];
        const fileElements = resourceElement.getChildren('file', imsNS);
        for (const fileElement of fileElements) {
          const fileHrefAttr = fileElement.getAttribute('href');
          if (fileHrefAttr) {
            files.push(fileHrefAttr.getValue());
          }
        }
        
        // Get title (simplified)
        let title = identifier;
        const titleElement = resourceElement.getChild('title', imsNS);
        if (titleElement) {
          title = titleElement.getText();
        }
        
        quizzes.push({
          identifier: identifier,
          href: href,
          files: files,
          title: title
        });
        
        console.log(`Added quiz: ${title} with ID: ${identifier}`);
      }
    }
    
    console.log(`Found ${quizzes.length} quizzes in manifest`);
    
    // For each quiz, find and parse the content
    for (let i = 0; i < quizzes.length; i++) {
      console.log(`Parsing quiz content for: ${quizzes[i].title}`);
      quizzes[i].content = findAndParseQuizContent(quizzes[i], tempFolder);
    }
    
    return quizzes;
  } catch (e) {
    console.log("Error processing manifest: " + e.message);
    return quizzes; // Return what we have so far
  }
}

/**
 * Finds and parses the content of a quiz.
 * 
 * @param {Object} quiz - Quiz information
 * @param {Folder} tempFolder - The folder containing extracted IMSCC
 * @return {Object} - Parsed quiz content
 */
function findAndParseQuizContent(quiz, tempFolder) {
  // Try to find the main quiz file
  let quizFile = null;
  
  try {
    // First check the href
    if (quiz.href) {
      quizFile = findFileByPath(tempFolder, quiz.href);
    }
    
    // If not found, check the files list
    if (!quizFile && quiz.files.length > 0) {
      for (const filePath of quiz.files) {
        if (filePath.endsWith('.xml')) {
          quizFile = findFileByPath(tempFolder, filePath);
          if (quizFile) {
            console.log(`Found quiz file: ${filePath}`);
            break;
          }
        }
      }
    }
    
    // Fallback - search for any XML files that might contain the quiz
    if (!quizFile) {
      console.log('Quiz file not found through standard paths. Attempting to search all extracted files...');
      const allFiles = getAllFilesInFolder(tempFolder);
      for (const file of allFiles) {
        if (file.getName().endsWith('.xml')) {
          const content = file.getBlob().getDataAsString();
          // Check if the file contains quiz-related content
          if (content.includes('questestinterop') || 
              content.includes('assessment') || 
              content.includes('qti') || 
              content.includes(quiz.identifier)) {
            console.log(`Found potential quiz file: ${file.getName()}`);
            quizFile = file;
            break;
          }
        }
      }
    }
    
    if (!quizFile) {
      console.log(`Could not find content for quiz ${quiz.identifier}`);
      return {
        title: quiz.title,
        description: '',
        questions: []
      };
    }
    
    // Parse the quiz XML content
    const content = quizFile.getBlob().getDataAsString();
    try {
      const document = XmlService.parse(content);
      return parseQuizXml(document, quiz);
    } catch (error) {
      console.log(`Error parsing quiz XML: ${error.message}`);
      return {
        title: quiz.title,
        description: '',
        questions: []
      };
    }
  } catch (e) {
    console.log("Error finding quiz content: " + e.message);
    return {
      title: quiz.title,
      description: '',
      questions: []
    };
  }
}

/**
 * Gets all files in a folder and its subfolders.
 * 
 * @param {Folder} folder - The folder to search
 * @return {Array} - Array of files
 */
function getAllFilesInFolder(folder) {
  let files = [];
  
  // Get files in this folder
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }
  
  // Get files in subfolders
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const subfolder = folderIterator.next();
    files = files.concat(getAllFilesInFolder(subfolder));
  }
  
  return files;
}

/**
 * Finds a file by its path within the temp folder structure.
 * 
 * @param {Folder} rootFolder - The root folder
 * @param {string} path - The path to the file
 * @return {File} - The file if found, otherwise null
 */
function findFileByPath(rootFolder, path) {
  if (!path) return null;
  
  const parts = path.split('/');
  let currentFolder = rootFolder;
  
  // Navigate through folders
  for (let i = 0; i < parts.length - 1; i++) {
    const folderName = parts[i];
    if (!folderName) continue; // Skip empty folder names
    
    const subFolders = currentFolder.getFoldersByName(folderName);
    if (subFolders.hasNext()) {
      currentFolder = subFolders.next();
    } else {
      return null; // Folder not found
    }
  }
  
  // Find the file
  const fileName = parts[parts.length - 1];
  if (!fileName) return null;
  
  const files = currentFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    return files.next();
  }
  
  return null; // File not found
}

/**
 * Parses quiz XML into a structured format.
 * 
 * @param {Document} document - The XML document
 * @param {Object} quiz - Quiz information
 * @return {Object} - Structured quiz data
 */
function parseQuizXml(document, quiz) {
  const root = document.getRootElement();
  const rootName = root.getName();
  
  console.log(`Quiz XML root element: ${rootName}`);
  
  // Handle different quiz formats
  if (rootName === 'questestinterop') {
    // QTI v1.2 format
    return parseQtiv1Quiz(root, quiz);
  } else if (rootName === 'assessment') {
    // QTI v2.1 format or similar
    return parseQtiv2Quiz(root, quiz);
  } else {
    console.log(`Unknown quiz format with root element: ${rootName}`);
    return {
      title: quiz.title,
      description: '',
      questions: []
    };
  }
}

/**
 * Parses QTI v1.2 quiz format.
 * 
 * @param {Element} root - The XML root element
 * @param {Object} quiz - Quiz information
 * @return {Object} - Structured quiz data
 */
function parseQtiv1Quiz(root, quiz) {
  const result = {
    title: quiz.title,
    description: '',
    questions: []
  };
  
  try {
    // Find assessment element
    const assessment = root.getChild('assessment');
    if (!assessment) {
      console.log('No assessment element found');
      return result;
    }
    
    // Extract quiz title if available
    const assessmentTitle = assessment.getAttribute('title');
    if (assessmentTitle) {
      result.title = assessmentTitle.getValue();
    }
    
    // Extract quiz description
    const metadata = assessment.getChild('qtimetadata');
    if (metadata) {
      const metadataFields = metadata.getChildren('qtimetadatafield');
      for (const field of metadataFields) {
        const label = field.getChildText('fieldlabel');
        const entry = field.getChildText('fieldentry');
        if (label === 'qmd_description' && entry) {
          result.description = entry;
        }
      }
    }
    
    console.log(`Processing QTI v1.2 quiz: ${result.title}`);
    
    // Extract quiz items/questions
    const items = assessment.getChildren('item');
    for (const item of items) {
      const question = parseQtiv1Item(item);
      if (question) {
        result.questions.push(question);
      }
    }
    
    // Extract sections which may contain items
    const sections = assessment.getChildren('section');
    for (const section of sections) {
      const sectionItems = section.getChildren('item');
      for (const item of sectionItems) {
        const question = parseQtiv1Item(item);
        if (question) {
          result.questions.push(question);
        }
      }
    }
    
    console.log(`Parsed ${result.questions.length} questions from QTI v1.2 quiz`);
    return result;
  } catch (e) {
    console.log("Error parsing QTI v1.2 quiz: " + e.message);
    return result;
  }
}

/**
 * Parses a QTI v1.2 item into a question object.
 * 
 * @param {Element} item - The item XML element
 * @return {Object} - Structured question data
 */
function parseQtiv1Item(item) {
  try {
    const identAttr = item.getAttribute('ident');
    const titleAttr = item.getAttribute('title');
    
    const question = {
      id: identAttr ? identAttr.getValue() : Utilities.getUuid(),
      title: titleAttr ? titleAttr.getValue() : '',
      type: '',
      text: '',
      choices: [],
      correctAnswer: null,
      points: 1
    };
    
    // Get question type
    const itemmetadata = item.getChild('itemmetadata');
    if (itemmetadata) {
      const qtimetadata = itemmetadata.getChild('qtimetadata');
      if (qtimetadata) {
        const metadatafields = qtimetadata.getChildren('qtimetadatafield');
        for (const field of metadatafields) {
          const label = field.getChildText('fieldlabel');
          const entry = field.getChildText('fieldentry');
          if (label === 'question_type' && entry) {
            question.type = entry;
          }
        }
      }
    }
    
    // Get question text from presentation
    const presentation = item.getChild('presentation');
    if (presentation) {
      const material = presentation.getChild('material');
      if (material) {
        const mattext = material.getChild('mattext');
        if (mattext) {
          question.text = mattext.getText() || '';
        }
      }
    }
    
    // Get choices from render_choice
    if (presentation) {
      const renderChoice = presentation.getChild('render_choice');
      if (renderChoice) {
        const responseLabels = renderChoice.getChildren('response_label');
        for (const responseLabel of responseLabels) {
          const identAttr = responseLabel.getAttribute('ident');
          if (!identAttr) continue;
          
          const ident = identAttr.getValue();
          const material = responseLabel.getChild('material');
          if (!material) continue;
          
          const mattext = material.getChild('mattext');
          if (!mattext) continue;
          
          const choiceText = mattext.getText() || '';
          
          question.choices.push({
            id: ident,
            text: choiceText
          });
        }
      }
    }
    
    // Get correct answer from resprocessing
    const resprocessing = item.getChild('resprocessing');
    if (resprocessing) {
      const respconditions = resprocessing.getChildren('respcondition');
      for (const respcondition of respconditions) {
        const conditionvar = respcondition.getChild('conditionvar');
        if (!conditionvar) continue;
        
        const varequal = conditionvar.getChild('varequal');
        if (varequal) {
          const correctAnswerId = varequal.getText();
          if (question.type === 'multiple_choice_question') {
            question.correctAnswer = correctAnswerId;
          } else if (question.type === 'multiple_answers_question') {
            if (!question.correctAnswer) question.correctAnswer = [];
            question.correctAnswer.push(correctAnswerId);
          }
        }
      }
    }
    
    return question;
  } catch (e) {
    console.log("Error parsing QTI v1.2 item: " + e.message);
    return null;
  }
}

/**
 * Parses QTI v2.1 quiz format.
 * 
 * @param {Element} root - The XML root element
 * @param {Object} quiz - Quiz information
 * @return {Object} - Structured quiz data
 */
function parseQtiv2Quiz(root, quiz) {
  const result = {
    title: quiz.title,
    description: '',
    questions: []
  };
  
  try {
    // Extract title from root element
    const titleAttr = root.getAttribute('title');
    if (titleAttr) {
      result.title = titleAttr.getValue();
    }
    
    console.log(`Processing QTI v2.1 quiz: ${result.title}`);
    
    // Extract questions
    const items = root.getChildren('item');
    for (const item of items) {
      const question = parseQtiv2Item(item);
      if (question) {
        result.questions.push(question);
      }
    }
    
    console.log(`Parsed ${result.questions.length} questions from QTI v2.1 quiz`);
    return result;
  } catch (e) {
    console.log("Error parsing QTI v2.1 quiz: " + e.message);
    return result;
  }
}

/**
 * Parses a QTI v2.1 item into a question object.
 * 
 * @param {Element} item - The item XML element
 * @return {Object} - Structured question data
 */
function parseQtiv2Item(item) {
  try {
    const identifierAttr = item.getAttribute('identifier');
    const titleAttr = item.getAttribute('title');
    
    const question = {
      id: identifierAttr ? identifierAttr.getValue() : Utilities.getUuid(),
      title: titleAttr ? titleAttr.getValue() : '',
      type: '',
      text: '',
      choices: [],
      correctAnswer: null,
      points: 1
    };
    
    // Get question body and type
    const itemBody = item.getChild('itemBody');
    if (itemBody) {
      // Determine question type based on interaction type
      if (itemBody.getChild('choiceInteraction')) {
        question.type = 'multiple_choice_question';
      } else if (itemBody.getChild('extendedTextInteraction')) {
        question.type = 'essay_question';
      } else if (itemBody.getChild('textEntryInteraction')) {
        question.type = 'short_answer_question';
      }
    
      // Get question text
      const promptElement = itemBody.getChild('prompt');
      if (promptElement) {
        question.text = promptElement.getText() || '';
      }
    
      // Get choices for multiple choice questions
      const choiceInteraction = itemBody.getChild('choiceInteraction');
      if (choiceInteraction) {
        const simpleChoices = choiceInteraction.getChildren('simpleChoice');
        for (const choice of simpleChoices) {
          const identifierAttr = choice.getAttribute('identifier');
          if (!identifierAttr) continue;
          
          const identifier = identifierAttr.getValue();
          const text = choice.getText() || '';
          
          question.choices.push({
            id: identifier,
            text: text
          });
        }
      }
    }
    
    // Get correct answer from responseDeclaration
    const responseDeclaration = item.getChild('responseDeclaration');
    if (responseDeclaration) {
      const correctResponse = responseDeclaration.getChild('correctResponse');
      if (correctResponse) {
        const values = correctResponse.getChildren('value');
        if (values.length === 1) {
          question.correctAnswer = values[0].getText();
        } else if (values.length > 1) {
          question.correctAnswer = values.map(v => v.getText());
        }
      }
    }
    
    return question;
  } catch (e) {
    console.log("Error parsing QTI v2.1 item: " + e.message);
    return null;
  }
}

/**
 * Creates Google Forms from parsed quiz data.
 * 
 * @param {Array} quizzes - Array of parsed quiz objects
 * @return {Array} - Array of created form IDs and URLs
 */
function createGoogleForms(quizzes) {
  const results = [];
  console.log(`Attempting to create forms for ${quizzes.length} parsed quiz objects.`);

  for (const quiz of quizzes) {
    // Log the quiz object being processed
    console.log(`Processing quiz: ${quiz.identifier}, Title hint: ${quiz.title}`);
    
    if (!quiz.content) {
      console.log(`Skipping quiz with no content: ${quiz.identifier} - ${quiz.title}`);
      continue;
    }

    // Log the content object
    console.log(`Quiz content found for ${quiz.identifier}: Title='${quiz.content.title}', Questions=${quiz.content.questions ? quiz.content.questions.length : 'N/A'}`);
    
    // Add a check for empty questions array specifically
    if (!quiz.content.questions || quiz.content.questions.length === 0) {
      console.log(`Skipping form creation for ${quiz.content.title} because it has no parsed questions.`);
      // Optionally, create an empty form anyway or handle as needed
      // continue; // Uncomment this line if you want to strictly skip forms with 0 questions
    }
    
    try {
      console.log(`Creating Google Form for: ${quiz.content.title}`);
      
      // Create a new form
      const form = FormApp.create(quiz.content.title || `Quiz ${quiz.identifier}`); // Add fallback title
      
      // Set description
      if (quiz.content.description) {
        form.setDescription(quiz.content.description);
      }
      
      // Add questions
      const questionCount = quiz.content.questions ? quiz.content.questions.length : 0;
      console.log(`Adding ${questionCount} questions to form: ${form.getTitle()}`);
      if (questionCount > 0) {
        for (const question of quiz.content.questions) {
          // Log each question before adding
          console.log(`Adding question: ID=${question.id}, Type=${question.type}, Text=${question.text ? question.text.substring(0, 50) + '...' : 'No Text'}`);
          addQuestionToForm(form, question);
        }
      } else {
        console.log(`No questions to add for form: ${form.getTitle()}`);
      }
      
      // Make it a quiz if needed (even if 0 questions, maybe?)
      // Check if *any* question had a correct answer defined during parsing
      const hasCorrectAnswers = quiz.content.questions && quiz.content.questions.some(qn => qn.correctAnswer !== null && qn.correctAnswer !== undefined);
      if (hasCorrectAnswers) {
        console.log(`Setting form ${form.getTitle()} as a quiz.`);
        form.setIsQuiz(true);
      } else {
         console.log(`Form ${form.getTitle()} will not be set as a quiz (no correct answers found).`);
      }
      
      // Log before pushing to results
      console.log(`Successfully processed form: ${form.getTitle()}. Adding to results.`);
      
      results.push({
        id: form.getId(),
        url: form.getPublishedUrl(),
        editUrl: form.getEditUrl(),
        title: form.getTitle(), // Use form's actual title
        numQuestions: questionCount,
        creationDate: new Date().toISOString()
      });
      
      console.log(`Form created: ${form.getEditUrl()}`);
      
    } catch (error) {
      // Log the specific error for this form
      console.log(`Error creating form for quiz "${quiz.content.title || quiz.identifier}": ${error.message}`);
      console.log(`Error stack: ${error.stack}`); // Log stack trace for more details
    }
  }
  
  console.log(`Finished createGoogleForms function. Total forms created and added to results: ${results.length}`);
  return results;
}

/**
 * Adds a question to a Google Form.
 * 
 * @param {Form} form - The Google Form
 * @param {Object} question - Question data
 */
function addQuestionToForm(form, question) {
  try {
    let formQuestion;
    
    // Default to something if text is empty
    if (!question.text || question.text.trim() === '') {
      console.log(`Question ID ${question.id} has empty text. Using title or default.`);
      question.text = question.title || `Question ${question.id}`;
    }
    
    // Log the final question text being used
    console.log(`Adding question with text: ${question.text.substring(0,100)}...`);

    // Handle different question types
    switch (question.type) {
      case 'multiple_choice_question':
        formQuestion = form.addMultipleChoiceItem();
        formQuestion.setTitle(question.text);
        
        if (question.choices && question.choices.length > 0) {
          const choices = question.choices.map(choice => {
            const isCorrect = choice.id === question.correctAnswer;
            return formQuestion.createChoice(choice.text, isCorrect);
          });
          formQuestion.setChoices(choices);
        } else {
          // Default choices if none are found
          formQuestion.setChoices([
            formQuestion.createChoice("Option 1"),
            formQuestion.createChoice("Option 2")
          ]);
        }
        break;
        
      case 'multiple_answers_question':
        formQuestion = form.addCheckboxItem();
        formQuestion.setTitle(question.text);
        
        if (question.choices && question.choices.length > 0) {
          const checkboxChoices = question.choices.map(choice => {
            const isCorrect = Array.isArray(question.correctAnswer) && 
                            question.correctAnswer.includes(choice.id);
            return formQuestion.createChoice(choice.text, isCorrect);
          });
          formQuestion.setChoices(checkboxChoices);
        } else {
          // Default choices if none are found
          formQuestion.setChoices([
            formQuestion.createChoice("Option 1"),
            formQuestion.createChoice("Option 2")
          ]);
        }
        break;
        
      case 'essay_question':
        formQuestion = form.addParagraphTextItem();
        formQuestion.setTitle(question.text);
        break;
        
      case 'short_answer_question':
        formQuestion = form.addTextItem();
        formQuestion.setTitle(question.text);
        break;
        
      case 'true_false_question':
        formQuestion = form.addMultipleChoiceItem();
        formQuestion.setTitle(question.text);
        
        const trueFalseChoices = [
          formQuestion.createChoice('True', question.correctAnswer === 'true'),
          formQuestion.createChoice('False', question.correctAnswer === 'false')
        ];
        formQuestion.setChoices(trueFalseChoices);
        break;
        
      default:
        // Default to paragraph text for unknown types
        console.log(`Unknown or unhandled question type: '${question.type}'. Defaulting to Paragraph Text for question: ${question.text.substring(0, 50)}...`);
        formQuestion = form.addParagraphTextItem();
        formQuestion.setTitle(question.text);
        break;
    }
    
    // Set points if it's a quiz
    if (form.isQuiz() && question.points > 0 && formQuestion && typeof formQuestion.setPoints === 'function') {
       console.log(`Setting points for question: ${question.points}`);
       formQuestion.setPoints(question.points);
    } else if (form.isQuiz() && (!formQuestion || typeof formQuestion.setPoints !== 'function')) {
       console.log(`Could not set points for question type ${question.type}`);
    }
  } catch (e) {
    console.log(`Error adding question (ID: ${question.id}, Text: ${question.text ? question.text.substring(0,50) : 'N/A'}...) to form: ${e.message}`);
    console.log(`Question add error stack: ${e.stack}`); // Log stack trace
  }
}

/**
 * Updates the active spreadsheet with form information.
 * 
 * @param {Array} formData - Array of form data objects
 */
function updateSpreadsheetWithFormInfo(formData) {
  try {
    // Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    
    // Check if spreadsheet is empty and add headers if needed
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(CONFIG.SPREADSHEET_HEADERS);
      sheet.getRange(1, 1, 1, CONFIG.SPREADSHEET_HEADERS.length).setFontWeight('bold');
      
      // Format the header row
      sheet.setFrozenRows(1);
    }
    
    // Add form data to the spreadsheet
    for (const form of formData) {
      const rowData = [
        form.title,
        form.url,
        form.editUrl,
        form.numQuestions,
        new Date(form.creationDate).toLocaleString()
      ];
      sheet.appendRow(rowData);
      
      // Format URL cells as hyperlinks
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 2).setFormula(`=HYPERLINK("${form.url}","View Form")`);
      sheet.getRange(lastRow, 3).setFormula(`=HYPERLINK("${form.editUrl}","Edit Form")`);
    }
    
    // Auto-resize columns to fit content
    for (let i = 1; i <= CONFIG.SPREADSHEET_HEADERS.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    console.log(`Added ${formData.length} forms to the spreadsheet`);
  } catch (e) {
    console.log(`Error updating spreadsheet: ${e.message}`);
  }
}

/**
 * Cleans up temporary folder after processing.
 * 
 * @param {Folder} tempFolder - The temporary folder
 */
function cleanupTempFolder(tempFolder) {
  try { // Add try-catch here for robustness
    if (tempFolder && typeof tempFolder.setTrashed === 'function') {
      tempFolder.setTrashed(true);
      console.log(`Temporary folder ${tempFolder.getName()} cleaned up.`);
    } else {
      console.log('Invalid tempFolder object provided for cleanup.');
    }
  } catch (e) {
    console.log(`Error cleaning up temp folder: ${e.message}`);
  }
}

/**
 * Alternative method to extract IMSCC file for cases where standard extraction fails.
 * This uses direct URL downloading which can sometimes work better with certain files.
 * 
 * @param {File} file - The IMSCC file from Drive
 * @param {Folder} tempFolder - The temporary folder
 * @return {Array} - Array of extracted files
 */
function alternativeExtractIMSCC(file, tempFolder) {
  console.log("Attempting alternative extraction method...");
  
  try {
    // Get download URL for the file
    const downloadUrl = file.getDownloadUrl();
    console.log(`File download URL acquired`);
    
    // Use URLFetchApp to download the file directly
    const response = UrlFetchApp.fetch(downloadUrl, {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to download file: HTTP ${response.getResponseCode()}`);
    }
    
    // Create a blob from the response content
    const fileBlob = response.getBlob().setContentType('application/zip');
    console.log(`Downloaded file as blob: ${fileBlob.getBytes().length} bytes`);
    
    // Attempt to unzip the blob
    const extracted = Utilities.unzip(fileBlob);
    console.log(`Unzipped ${extracted.length} files/folders`);
    
    // Process the extracted files
    const extractedFiles = [];
    
    for (let i = 0; i < extracted.length; i++) {
      const extractedBlob = extracted[i];
      
      // Skip directories
      if (extractedBlob.getName().endsWith('/')) {
        continue;
      }
      
      // Create the file in the temp folder (simplified approach)
      try {
        const newFile = tempFolder.createFile(extractedBlob);
        extractedFiles.push(newFile);
      } catch (e) {
        console.log(`Error creating file ${extractedBlob.getName()}: ${e.message}`);
      }
    }
    
    return extractedFiles;
  } catch (error) {
    console.log(`Alternative extraction failed: ${error.message}`);
    throw error;
  }
}