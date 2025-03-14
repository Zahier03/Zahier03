// Visual Scripting IDE for Google Apps Script
// Main Server-Side Code

/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Visual Script Editor')
      .addItem('Open Editor', 'openEditor')
      .addToUi();
}

/**
 * Opens the web app in a dialog.
 */
function openEditor() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(1200)
      .setHeight(800)
      .setTitle('Visual Script Editor');
  SpreadsheetApp.getUi().showModalDialog(html, 'Visual Script Editor');
}

/**
 * Serves the HTML content for the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Visual Script Editor')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns a list of available Google Apps Script functions and methods
 * to populate the block library.
 */
function getAvailableFunctions() {
  // This is just a sample of common Apps Script functions
  // You would expand this based on your needs
  return {
    spreadsheet: {
      category: "Spreadsheet",
      functions: [
        {name: "getActiveSpreadsheet", params: [], description: "Gets the active spreadsheet"},
        {name: "getActiveSheet", params: [], description: "Gets the active sheet in the active spreadsheet"},
        {name: "getRange", params: ["row", "column", "numRows", "numColumns"], description: "Gets a range at the specified coordinates"},
        {name: "getValue", params: [], description: "Gets the value of a range or cell"},
        {name: "setValue", params: ["value"], description: "Sets the value of a range or cell"},
        {name: "getValues", params: [], description: "Gets the values of a range as a 2D array"},
        {name: "setValues", params: ["values"], description: "Sets the values of a range from a 2D array"}
      ]
    },
    ui: {
      category: "User Interface",
      functions: [
        {name: "alert", params: ["message"], description: "Shows an alert dialog"},
        {name: "prompt", params: ["message", "title"], description: "Shows a prompt dialog"},
        {name: "confirm", params: ["message"], description: "Shows a confirmation dialog"}
      ]
    },
    utilities: {
      category: "Utilities",
      functions: [
        {name: "sleep", params: ["milliseconds"], description: "Suspends execution for the specified duration"},
        {name: "formatDate", params: ["date", "timeZone", "format"], description: "Formats a date according to the pattern"}
      ]
    },
    flow: {
      category: "Control Flow",
      functions: [
        {name: "if", params: ["condition", "thenBlock", "elseBlock"], description: "Conditional execution"},
        {name: "forEach", params: ["array", "loopBlock"], description: "Loop through elements in an array"},
        {name: "while", params: ["condition", "loopBlock"], description: "Loop while a condition is true"},
        {name: "function", params: ["name", "params", "functionBlock"], description: "Define a function"}
      ]
    }
  };
}

/**
 * Generates Apps Script code from the block structure.
 * @param {Object} blockStructure - The structure of connected blocks
 * @return {String} Generated Apps Script code
 */
function generateCode(blockStructure) {
  // This is a simplified version - a real implementation would be more complex
  let code = "// Generated Apps Script Code\n\n";
  
  // Process each function in the block structure
  for (let funcId in blockStructure.functions) {
    const func = blockStructure.functions[funcId];
    code += `function ${func.name}(${func.parameters.join(', ')}) {\n`;
    
    // Process blocks inside the function
    code += processBlocks(func.blocks, 2);
    
    code += "}\n\n";
  }
  
  return code;
}

/**
 * Process blocks and convert them to code
 * @param {Array} blocks - Array of block objects
 * @param {Number} indent - Indentation level
 * @return {String} Generated code for the blocks
 */
function processBlocks(blocks, indent) {
  let code = "";
  const indentStr = " ".repeat(indent);
  
  blocks.forEach(block => {
    switch (block.type) {
      case "spreadsheet.getActiveSpreadsheet":
        code += `${indentStr}const ss = SpreadsheetApp.getActiveSpreadsheet();\n`;
        break;
      case "spreadsheet.getActiveSheet":
        code += `${indentStr}const sheet = ${block.parent || "SpreadsheetApp"}.getActiveSheet();\n`;
        break;
      case "spreadsheet.getRange":
        const params = [block.row, block.column, block.numRows, block.numColumns].filter(p => p !== undefined);
        code += `${indentStr}const range = ${block.parent || "sheet"}.getRange(${params.join(', ')});\n`;
        break;
      case "spreadsheet.getValue":
        code += `${indentStr}const value = ${block.parent || "range"}.getValue();\n`;
        break;
      case "spreadsheet.setValue":
        code += `${indentStr}${block.parent || "range"}.setValue(${block.value});\n`;
        break;
      case "ui.alert":
        code += `${indentStr}SpreadsheetApp.getUi().alert("${block.message}");\n`;
        break;
      case "flow.if":
        code += `${indentStr}if (${block.condition}) {\n`;
        if (block.thenBlocks) {
          code += processBlocks(block.thenBlocks, indent + 2);
        }
        code += `${indentStr}} else {\n`;
        if (block.elseBlocks) {
          code += processBlocks(block.elseBlocks, indent + 2);
        }
        code += `${indentStr}}\n`;
        break;
      case "flow.forEach":
        code += `${indentStr}${block.array}.forEach(function(item) {\n`;
        if (block.loopBlocks) {
          code += processBlocks(block.loopBlocks, indent + 2);
        }
        code += `${indentStr}});\n`;
        break;
      default:
        // Generic function call
        if (block.type.includes('.')) {
          const [obj, method] = block.type.split('.');
          const paramStr = (block.params || []).join(', ');
          code += `${indentStr}${block.variableName ? `const ${block.variableName} = ` : ''}${block.parent || obj}.${method}(${paramStr});\n`;
        } else {
          code += `${indentStr}// Unknown block type: ${block.type}\n`;
        }
    }
  });
  
  return code;
}

/**
 * Saves the generated code as a new Apps Script project
 * @param {String} code - The generated Apps Script code
 * @param {String} projectName - Name for the new project
 * @return {Object} Result of the save operation
 */
function saveAsNewProject(code, projectName) {
  try {
    // In a real implementation, you would use the Apps Script API to create a new project
    // This is a simplified version that requires manual copy-paste
    return {
      success: true,
      message: "Code generated successfully. Use the 'Export Code' button to download and create a new Apps Script project manually.",
      code: code
    };
  } catch (e) {
    return {
      success: false,
      message: "Error creating project: " + e.toString()
    };
  }
}

/**
 * Gets a list of HTML components that can be dragged onto the canvas
 */
function getHtmlComponents() {
  return [
    {type: "div", name: "Container", properties: {width: "100px", height: "100px", backgroundColor: "#f0f0f0", color: "#000000", padding: "10px"}},
    {type: "button", name: "Button", properties: {text: "Click Me", backgroundColor: "#4CAF50", color: "white", padding: "10px", borderRadius: "4px"}},
    {type: "input", name: "Text Input", properties: {placeholder: "Enter text...", width: "150px", padding: "8px"}},
    {type: "select", name: "Dropdown", properties: {options: ["Option 1", "Option 2", "Option 3"], width: "150px", padding: "8px"}},
    {type: "table", name: "Table", properties: {rows: 3, columns: 3, width: "200px", borderCollapse: "collapse"}},
    {type: "label", name: "Label", properties: {text: "Text Label", fontWeight: "bold"}}
  ];
}
