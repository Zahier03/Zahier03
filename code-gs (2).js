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
    variables: {
      category: "Variables",
      functions: [
        {name: "declareVariable", params: ["name", "value"], description: "Declare a variable with a value", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "getVariable", params: ["name"], description: "Get a variable's value", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "setVariable", params: ["name", "value"], description: "Set a variable's value", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "createArray", params: ["elements"], description: "Create an array with elements", returnValue: true, outputShape: "puzzle-out", inputShape: "none"}
      ]
    },
    spreadsheet: {
      category: "Spreadsheet",
      functions: [
        {name: "getActiveSpreadsheet", params: [], description: "Gets the active spreadsheet", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "getActiveSheet", params: [], description: "Gets the active sheet in the active spreadsheet", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "getRange", params: ["row", "column", "numRows", "numColumns"], description: "Gets a range at the specified coordinates", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "getValue", params: [], description: "Gets the value of a range or cell", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "setValue", params: ["value"], description: "Sets the value of a range or cell", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "getValues", params: [], description: "Gets the values of a range as a 2D array", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "setValues", params: ["values"], description: "Sets the values of a range from a 2D array", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"}
      ]
    },
    ui: {
      category: "User Interface",
      functions: [
        {name: "alert", params: ["message"], description: "Shows an alert dialog", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "prompt", params: ["message", "title"], description: "Shows a prompt dialog", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "confirm", params: ["message"], description: "Shows a confirmation dialog", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "createHtmlOutput", params: ["html"], description: "Creates HTML output", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "showModalDialog", params: ["html", "title"], description: "Shows modal dialog with HTML content", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"}
      ]
    },
    utilities: {
      category: "Utilities",
      functions: [
        {name: "sleep", params: ["milliseconds"], description: "Suspends execution for the specified duration", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "formatDate", params: ["date", "timeZone", "format"], description: "Formats a date according to the pattern", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "parseCsv", params: ["csv"], description: "Parses CSV data into a 2D array", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "base64Encode", params: ["data"], description: "Encodes data as Base64", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "base64Decode", params: ["encoded"], description: "Decodes Base64 data", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"}
      ]
    },
    display: {
      category: "Display",
      functions: [
        {name: "logOutput", params: ["message"], description: "Logs a message to the console", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "showResult", params: ["title", "message"], description: "Shows a result dialog with a title and message", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "createChart", params: ["title", "data", "type"], description: "Creates a chart with the given data", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "displayTable", params: ["data", "headers"], description: "Displays data as a table", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "createDashboard", params: ["title", "components"], description: "Creates a dashboard with components", returnValue: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "simpleOutput", params: ["value"], description: "Simple output display of a value", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "appendToOutput", params: ["value"], description: "Append value to output log", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "clearOutput", params: [], description: "Clear the output display", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in"}
      ]
    },
    flow: {
      category: "Control Flow",
      functions: [
        {name: "if", params: ["condition"], description: "Conditional execution", isContainer: true, hasElse: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "forEach", params: ["array", "itemName"], description: "Loop through elements in an array", isContainer: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "for", params: ["start", "end", "step", "counterName"], description: "Loop from start to end", isContainer: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "while", params: ["condition"], description: "Loop while a condition is true", isContainer: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "doWhile", params: ["condition"], description: "Loop at least once, then while condition is true", isContainer: true, outputShape: "puzzle-out", inputShape: "puzzle-in"},
        {name: "function", params: ["name", "params"], description: "Define a function", isContainer: true, isFunctionDef: true, outputShape: "none", inputShape: "none"}
      ]
    },
    logic: {
      category: "Logic",
      functions: [
        {name: "and", params: ["left", "right"], description: "Logical AND of two conditions", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "or", params: ["left", "right"], description: "Logical OR of two conditions", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "not", params: ["condition"], description: "Logical NOT of a condition", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "equals", params: ["left", "right"], description: "Equality comparison", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "notEquals", params: ["left", "right"], description: "Inequality comparison", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "greaterThan", params: ["left", "right"], description: "Greater than comparison", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "lessThan", params: ["left", "right"], description: "Less than comparison", returnValue: true, outputShape: "puzzle-out", inputShape: "none"}
      ]
    },
    math: {
      category: "Math",
      functions: [
        {name: "calculate", params: ["expression"], description: "Calculate a mathematical expression", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "add", params: ["left", "right"], description: "Addition", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "subtract", params: ["left", "right"], description: "Subtraction", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "multiply", params: ["left", "right"], description: "Multiplication", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "divide", params: ["left", "right"], description: "Division", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "mod", params: ["left", "right"], description: "Modulus (remainder)", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "power", params: ["base", "exponent"], description: "Raise a number to a power", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "sqrt", params: ["number"], description: "Square root of a number", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "random", params: ["min", "max"], description: "Random number between min and max", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "round", params: ["number"], description: "Round to nearest integer", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "floor", params: ["number"], description: "Round down to nearest integer", returnValue: true, outputShape: "puzzle-out", inputShape: "none"},
        {name: "ceil", params: ["number"], description: "Round up to nearest integer", returnValue: true, outputShape: "puzzle-out", inputShape: "none"}
      ]
    },
    operator: {
      category: "Operators",
      functions: [
        {name: "plus", params: ["left", "right"], description: "Addition (+)", returnValue: true, outputShape: "puzzle-out", inputShape: "none", symbol: "+"},
        {name: "minus", params: ["left", "right"], description: "Subtraction (-)", returnValue: true, outputShape: "puzzle-out", inputShape: "none", symbol: "-"},
        {name: "times", params: ["left", "right"], description: "Multiplication (*)", returnValue: true, outputShape: "puzzle-out", inputShape: "none", symbol: "*"},
        {name: "dividedBy", params: ["left", "right"], description: "Division (/)", returnValue: true, outputShape: "puzzle-out", inputShape: "none", symbol: "/"},
        {name: "modulo", params: ["left", "right"], description: "Modulus (%)", returnValue: true, outputShape: "puzzle-out", inputShape: "none", symbol: "%"},
        {name: "increment", params: ["variable"], description: "Increment (++)", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in", inlineCode: true},
        {name: "decrement", params: ["variable"], description: "Decrement (--)", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in", inlineCode: true},
        {name: "assign", params: ["variable", "value"], description: "Assignment (=)", returnValue: false, outputShape: "puzzle-out", inputShape: "puzzle-in", inlineCode: true}
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
  let code = "// Generated Apps Script Code\n\n";
  let variablesCode = "";
  let functionsCode = "";
  let mainCode = "";
  
  // Track defined variables for scoping
  const definedVariables = new Set();
  
  // Process each function in the block structure
  for (let funcId in blockStructure.functions) {
    const func = blockStructure.functions[funcId];
    
    // Skip the main function for now - we'll add it last
    if (func.name === 'main') {
      mainCode = `function ${func.name}() {\n${processBlocks(func.blocks, 2, definedVariables)}\n}\n\n`;
      continue;
    }
    
    // Clear defined variables for each function (new scope)
    definedVariables.clear();
    
    // Add function parameters to defined variables
    if (func.parameters) {
      func.parameters.forEach(param => definedVariables.add(param));
    }
    
    // Generate function code
    functionsCode += `function ${func.name}(${(func.parameters || []).join(', ')}) {\n`;
    functionsCode += processBlocks(func.blocks, 2, definedVariables);
    functionsCode += "}\n\n";
  }
  
  // Add HTML component code if any
  if (blockStructure.htmlComponents && blockStructure.htmlComponents.length > 0) {
    code += "// HTML Components\n";
    blockStructure.htmlComponents.forEach(comp => {
      code += `const ${comp.id} = document.createElement("${comp.type}");\n`;
      if (comp.properties) {
        for (let prop in comp.properties) {
          code += `${comp.id}.style.${prop} = "${comp.properties[prop]}";\n`;
        }
      }
      code += "\n";
    });
  }
  
  // Combine all code sections
  code += variablesCode + functionsCode + mainCode;
  
  // Add trigger function for HTML UI if needed
  if (blockStructure.hasHtmlComponents) {
    code += `
/**
 * Creates UI elements and shows the web app
 */
function doGet() {
  const htmlOutput = HtmlService.createHtmlOutput();
  
  // Add HTML content
  htmlOutput.setContent(\`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Generated App</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
        </style>
      </head>
      <body>
        <div id="app-container"></div>
        <script>
          // Initialization code
          document.addEventListener('DOMContentLoaded', function() {
            google.script.run
              .withSuccessHandler(function(result) {
                console.log('App initialized');
              })
              .main();
          });
        </script>
      </body>
    </html>
  \`);
  
  return htmlOutput.setTitle('Generated App');
}
`;
  }
  
  return code;
}

/**
 * Process blocks and convert them to code
 * @param {Array} blocks - Array of block objects
 * @param {Number} indent - Indentation level
 * @param {Set} definedVariables - Set of variables defined in current scope
 * @return {String} Generated code for the blocks
 */
function processBlocks(blocks, indent, definedVariables) {
  if (!blocks || !blocks.length) return "";
  
  let code = "";
  const indentStr = " ".repeat(indent);
  
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const nextBlock = i < blocks.length - 1 ? blocks[i + 1] : null;
    
    // Skip if block is not defined
    if (!block || !block.type) continue;
    
    // Split type into category and function
    const [category, method] = block.type.split('.');
    
    // Process based on block category and method
    switch (category) {
      case "variables":
        // Process variable blocks
        switch (method) {
          case "declareVariable":
            const varName = block.params?.name || "myVar";
            const varValue = block.params?.value || '""';
            if (!definedVariables.has(varName)) {
              code += `${indentStr}let ${varName} = ${varValue};\n`;
              definedVariables.add(varName);
            } else {
              code += `${indentStr}${varName} = ${varValue};\n`;
            }
            break;
          case "getVariable":
            // Just reference the variable - typically used in expressions
            // No code generated here directly
            break;
          case "setVariable":
            const setVarName = block.params?.name || "myVar";
            const setValue = block.params?.value || '""';
            code += `${indentStr}${setVarName} = ${setValue};\n`;
            break;
          case "createArray":
            const arrayName = block.params?.name || "myArray";
            const elements = block.params?.elements || "[]";
            if (!definedVariables.has(arrayName)) {
              code += `${indentStr}let ${arrayName} = ${elements};\n`;
              definedVariables.add(arrayName);
            } else {
              code += `${indentStr}${arrayName} = ${elements};\n`;
            }
            break;
        }
        break;
        
      case "spreadsheet":
        // Process spreadsheet blocks
        switch (method) {
          case "getActiveSpreadsheet":
            const ssVarName = block.variableName || "ss";
            code += `${indentStr}const ${ssVarName} = SpreadsheetApp.getActiveSpreadsheet();\n`;
            definedVariables.add(ssVarName);
            break;
          case "getActiveSheet":
            const sheetVarName = block.variableName || "sheet";
            const ssRef = block.sourceRef || "SpreadsheetApp";
            code += `${indentStr}const ${sheetVarName} = ${ssRef}.getActiveSheet();\n`;
            definedVariables.add(sheetVarName);
            break;
          case "getRange":
            const rangeVarName = block.variableName || "range";
            const sheetRef = block.sourceRef || "sheet";
            const row = block.params?.row || "1";
            const col = block.params?.column || "1";
            const rows = block.params?.numRows || "1";
            const cols = block.params?.numColumns || "1";
            code += `${indentStr}const ${rangeVarName} = ${sheetRef}.getRange(${row}, ${col}, ${rows}, ${cols});\n`;
            definedVariables.add(rangeVarName);
            break;
          case "getValue":
            const valueVarName = block.variableName || "value";
            const rangeRef = block.sourceRef || "range";
            code += `${indentStr}const ${valueVarName} = ${rangeRef}.getValue();\n`;
            definedVariables.add(valueVarName);
            break;
          case "setValue":
            const setRangeRef = block.sourceRef || "range";
            const setValue2 = block.params?.value || '""';
            code += `${indentStr}${setRangeRef}.setValue(${setValue2});\n`;
            break;
          case "getValues":
            const valuesVarName = block.variableName || "values";
            const rangeRef2 = block.sourceRef || "range";
            code += `${indentStr}const ${valuesVarName} = ${rangeRef2}.getValues();\n`;
            definedVariables.add(valuesVarName);
            break;
          case "setValues":
            const setRangeRef2 = block.sourceRef || "range";
            const setValues = block.params?.values || "[[]]";
            code += `${indentStr}${setRangeRef2}.setValues(${setValues});\n`;
            break;
        }
        break;
        
      case "ui":
        // Process UI blocks
        switch (method) {
          case "alert":
            const message = block.params?.message || '""';
            code += `${indentStr}SpreadsheetApp.getUi().alert(${message});\n`;
            break;
          case "prompt":
            const promptMsg = block.params?.message || '""';
            const promptTitle = block.params?.title || '""';
            const promptVarName = block.variableName || "result";
            code += `${indentStr}const ${promptVarName} = SpreadsheetApp.getUi().prompt(${promptMsg}, ${promptTitle});\n`;
            definedVariables.add(promptVarName);
            break;
          case "confirm":
            const confirmMsg = block.params?.message || '""';
            const confirmVarName = block.variableName || "result";
            code += `${indentStr}const ${confirmVarName} = SpreadsheetApp.getUi().alert(${confirmMsg}, SpreadsheetApp.getUi().ButtonSet.YES_NO) === SpreadsheetApp.getUi().Button.YES;\n`;
            definedVariables.add(confirmVarName);
            break;
          case "createHtmlOutput":
            const htmlContent = block.params?.html || '""';
            const htmlVarName = block.variableName || "htmlOutput";
            code += `${indentStr}const ${htmlVarName} = HtmlService.createHtmlOutput(${htmlContent});\n`;
            definedVariables.add(htmlVarName);
            break;
          case "showModalDialog":
            const htmlRef = block.params?.html || "htmlOutput";
            const dialogTitle = block.params?.title || '""';
            code += `${indentStr}SpreadsheetApp.getUi().showModalDialog(${htmlRef}, ${dialogTitle});\n`;
            break;
        }
        break;
        
      case "flow":
        // Process control flow blocks
        switch (method) {
          case "if":
            const condition = block.params?.condition || "true";
            code += `${indentStr}if (${condition}) {\n`;
            if (block.childBlocks) {
              code += processBlocks(block.childBlocks, indent + 2, new Set([...definedVariables]));
            }
            code += `${indentStr}} `;
            if (block.hasElse) {
              code += `else {\n`;
              if (block.elseBlocks) {
                code += processBlocks(block.elseBlocks, indent + 2, new Set([...definedVariables]));
              }
              code += `${indentStr}}`;
            }
            code += "\n";
            break;
          case "forEach":
            const array = block.params?.array || "[]";
            const itemName = block.params?.itemName || "item";
            code += `${indentStr}${array}.forEach(function(${itemName}) {\n`;
            // Add item to defined variables inside the loop
            const loopVars = new Set([...definedVariables]);
            loopVars.add(itemName);
            if (block.childBlocks) {
              code += processBlocks(block.childBlocks, indent + 2, loopVars);
            }
            code += `${indentStr}});\n`;
            break;
          case "for":
            const start = block.params?.start || "0";
            const end = block.params?.end || "10";
            const step = block.params?.step || "1";
            const counterName = block.params?.counterName || "i";
            code += `${indentStr}for (let ${counterName} = ${start}; ${counterName} < ${end}; ${counterName} += ${step}) {\n`;
            // Add counter to defined variables inside the loop
            const forLoopVars = new Set([...definedVariables]);
            forLoopVars.add(counterName);
            if (block.childBlocks) {
              code += processBlocks(block.childBlocks, indent + 2, forLoopVars);
            }
            code += `${indentStr}}\n`;
            break;
          case "while":
            const whileCondition = block.params?.condition || "true";
            code += `${indentStr}while (${whileCondition}) {\n`;
            if (block.childBlocks) {
              code += processBlocks(block.childBlocks, indent + 2, new Set([...definedVariables]));
            }
            code += `${indentStr}}\n`;
            break;
          case "doWhile":
            const doWhileCondition = block.params?.condition || "true";
            code += `${indentStr}do {\n`;
            if (block.childBlocks) {
              code += processBlocks(block.childBlocks, indent + 2, new Set([...definedVariables]));
            }
            code += `${indentStr}} while (${doWhileCondition});\n`;
            break;
          case "function":
            // Functions are processed separately, not inline
            break;
        }
        break;
        
      case "logic":
        // Logic operations (used as expressions, not statements)
        // These typically don't generate standalone code
        break;
        
      case "math":
        // Process math blocks
        switch (method) {
          case "calculate":
            const expr = block.params?.expression || "0";
            const calcVarName = block.variableName || "result";
            code += `${indentStr}const ${calcVarName} = ${expr};\n`;
            definedVariables.add(calcVarName);
            break;
          case "add":
            const leftAdd = block.params?.left || "0";
            const rightAdd = block.params?.right || "0";
            const addVarName = block.variableName || "sum";
            code += `${indentStr}const ${addVarName} = Number(${leftAdd}) + Number(${rightAdd});\n`;
            definedVariables.add(addVarName);
            break;
          case "subtract":
            const leftSub = block.params?.left || "0";
            const rightSub = block.params?.right || "0";
            const subVarName = block.variableName || "difference";
            code += `${indentStr}const ${subVarName} = ${leftSub} - ${rightSub};\n`;
            definedVariables.add(subVarName);
            break;
          case "multiply":
            const leftMult = block.params?.left || "0";
            const rightMult = block.params?.right || "0";
            const multVarName = block.variableName || "product";
            code += `${indentStr}const ${multVarName} = ${leftMult} * ${rightMult};\n`;
            definedVariables.add(multVarName);
            break;
          case "divide":
            const leftDiv = block.params?.left || "0";
            const rightDiv = block.params?.right || "1";
            const divVarName = block.variableName || "quotient";
            code += `${indentStr}const ${divVarName} = ${leftDiv} / ${rightDiv};\n`;
            definedVariables.add(divVarName);
            break;
          case "power":
            const base = block.params?.base || "0";
            const exponent = block.params?.exponent || "1";
            const powerVarName = block.variableName || "power";
            code += `${indentStr}const ${powerVarName} = Math.pow(${base}, ${exponent});\n`;
            definedVariables.add(powerVarName);
            break;
          case "sqrt":
            const sqrtNum = block.params?.number || "0";
            const sqrtVarName = block.variableName || "squareRoot";
            code += `${indentStr}const ${sqrtVarName} = Math.sqrt(${sqrtNum});\n`;
            definedVariables.add(sqrtVarName);
            break;
          default:
            // Other math operations
            if (block.variableName) {
              code += `${indentStr}const ${block.variableName} = Math.${method}(${Object.values(block.params || {}).join(', ')});\n`;
              definedVariables.add(block.variableName);
            }
        }
        break;
        
      case "operator":
        // Process operator blocks
        switch (method) {
          case "plus":
            const leftPlus = block.params?.left || "0";
            const rightPlus = block.params?.right || "0";
            const plusVarName = block.variableName || "result";
            code += `${indentStr}const ${plusVarName} = ${leftPlus} + ${rightPlus};\n`;
            definedVariables.add(plusVarName);
            break;
          case "minus":
            const leftMinus = block.params?.left || "0";
            const rightMinus = block.params?.right || "0";
            const minusVarName = block.variableName || "result";
            code += `${indentStr}const ${minusVarName} = ${leftMinus} - ${rightMinus};\n`;
            definedVariables.add(minusVarName);
            break;
          case "times":
            const leftTimes = block.params?.left || "0";
            const rightTimes = block.params?.right || "0";
            const timesVarName = block.variableName || "result";
            code += `${indentStr}const ${timesVarName} = ${leftTimes} * ${rightTimes};\n`;
            definedVariables.add(timesVarName);
            break;
          case "dividedBy":
            const leftDivBy = block.params?.left || "0";
            const rightDivBy = block.params?.right || "1";
            const divByVarName = block.variableName || "result";
            code += `${indentStr}const ${divByVarName} = ${leftDivBy} / ${rightDivBy};\n`;
            definedVariables.add(divByVarName);
            break;
          case "modulo":
            const leftMod = block.params?.left || "0";
            const rightMod = block.params?.right || "1";
            const modVarName = block.variableName || "result";
            code += `${indentStr}const ${modVarName} = ${leftMod} % ${rightMod};\n`;
            definedVariables.add(modVarName);
            break;
          case "increment":
            const incVar = block.params?.variable || "x";
            code += `${indentStr}${incVar}++;\n`;
            break;
          case "decrement":
            const decVar = block.params?.variable || "x";
            code += `${indentStr}${decVar}--;\n`;
            break;
          case "assign":
            const assignVar = block.params?.variable || "x";
            const assignVal = block.params?.value || "0";
            code += `${indentStr}${assignVar} = ${assignVal};\n`;
            break;
        }
        break;
        
      case "display":
        // Process display blocks
        switch (method) {
          case "logOutput":
            const logMessage = block.params?.message || '""';
            code += `${indentStr}console.log(${logMessage});\n`;
            break;
          case "showResult":
            const resultTitle = block.params?.title || '"Result"';
            const resultMessage = block.params?.message || '""';
            code += `${indentStr}SpreadsheetApp.getUi().alert(${resultTitle}, ${resultMessage}, SpreadsheetApp.getUi().ButtonSet.OK);\n`;
            break;
          case "simpleOutput":
            const outputValue = block.params?.value || '""';
            code += `${indentStr}console.log("Output: " + (${outputValue}));\n` +
                   `${indentStr}// For UI display\n` +
                   `${indentStr}if (typeof outputElement !== 'undefined') {\n` +
                   `${indentStr}  outputElement.innerHTML += "<div>Output: " + (${outputValue}) + "</div>";\n` +
                   `${indentStr}}\n`;
            break;
          case "appendToOutput":
            const appendValue = block.params?.value || '""';
            code += `${indentStr}console.log(${appendValue});\n` +
                   `${indentStr}// For UI display\n` +
                   `${indentStr}if (typeof outputElement !== 'undefined') {\n` +
                   `${indentStr}  outputElement.innerHTML += "<div>" + (${appendValue}) + "</div>";\n` +
                   `${indentStr}}\n`;
            break;
          case "clearOutput":
            code += `${indentStr}// Clear the output display\n` +
                   `${indentStr}if (typeof outputElement !== 'undefined') {\n` +
                   `${indentStr}  outputElement.innerHTML = "";\n` +
                   `${indentStr}}\n`;
            break;
          case "createChart":
            const chartTitle = block.params?.title || '""';
            const chartData = block.params?.data || "[]";
            const chartType = block.params?.type || '"LINE"';
            const chartVarName = block.variableName || "chart";
            code += `${indentStr}const ${chartVarName} = Charts.newLineChart()\n` +
                   `${indentStr}  .setTitle(${chartTitle})\n` +
                   `${indentStr}  .setDataTable(${chartData})\n` +
                   `${indentStr}  .build();\n`;
            definedVariables.add(chartVarName);
            break;
          case "displayTable":
            const tableData = block.params?.data || "[]";
            const tableHeaders = block.params?.headers || "[]";
            const tableVarName = block.variableName || "table";
            code += `${indentStr}// Create HTML table from data\n` +
                   `${indentStr}let ${tableVarName} = '<table border=\"1\" style=\"border-collapse: collapse;\">';\n` +
                   `${indentStr}// Add headers\n` +
                   `${indentStr}${tableVarName} += '<tr>';\n` +
                   `${indentStr}for (const header of ${tableHeaders}) {\n` +
                   `${indentStr}  ${tableVarName} += '<th>' + header + '</th>';\n` +
                   `${indentStr}}\n` +
                   `${indentStr}${tableVarName} += '</tr>';\n` +
                   `${indentStr}// Add data rows\n` +
                   `${indentStr}for (const row of ${tableData}) {\n` +
                   `${indentStr}  ${tableVarName} += '<tr>';\n` +
                   `${indentStr}  for (const cell of row) {\n` +
                   `${indentStr}    ${tableVarName} += '<td>' + cell + '</td>';\n` +
                   `${indentStr}  }\n` +
                   `${indentStr}  ${tableVarName} += '</tr>';\n` +
                   `${indentStr}}\n` +
                   `${indentStr}${tableVarName} += '</table>';\n`;
            definedVariables.add(tableVarName);
            break;
        }
        break;
        
      case "utilities":
        // Process utility blocks
        switch (method) {
          case "sleep":
            const milliseconds = block.params?.milliseconds || "1000";
            code += `${indentStr}Utilities.sleep(${milliseconds});\n`;
            break;
          case "formatDate":
            const date = block.params?.date || "new Date()";
            const timeZone = block.params?.timeZone || "Session.getScriptTimeZone()";
            const format = block.params?.format || '"yyyy-MM-dd"';
            const dateVarName = block.variableName || "formattedDate";
            code += `${indentStr}const ${dateVarName} = Utilities.formatDate(${date}, ${timeZone}, ${format});\n`;
            definedVariables.add(dateVarName);
            break;
          case "parseCsv":
            const csv = block.params?.csv || '""';
            const csvVarName = block.variableName || "parsedCsv";
            code += `${indentStr}const ${csvVarName} = Utilities.parseCsv(${csv});\n`;
            definedVariables.add(csvVarName);
            break;
          case "base64Encode":
            const data = block.params?.data || '""';
            const b64EncVarName = block.variableName || "encodedData";
            code += `${indentStr}const ${b64EncVarName} = Utilities.base64Encode(${data});\n`;
            definedVariables.add(b64EncVarName);
            break;
          case "base64Decode":
            const encoded = block.params?.encoded || '""';
            const b64DecVarName = block.variableName || "decodedData";
            code += `${indentStr}const ${b64DecVarName} = Utilities.base64Decode(${encoded});\n`;
            definedVariables.add(b64DecVarName);
            break;
        }
        break;
        
      default:
        // Generic function call for unknown categories
        code += `${indentStr}// ${block.type}\n`;
        break;
    }
  }
  
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
    // This is a simplified version that provides a downloadable code file
    
    // Create code content with HTML wrapper for download
    const htmlCode = `<!DOCTYPE html>
<html>
<head>
  <title>${projectName || 'Generated Apps Script'}</title>
  <script>
    function downloadCode() {
      const blob = new Blob([document.getElementById('codeContent').textContent], {type: 'text/javascript'});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Code.gs';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  </script>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    pre { background-color: #f5f5f5; padding: 15px; border-radius: 5px; overflow: auto; }
    .button { background-color: #4CAF50; color: white; padding: 10px 20px; border: none; 
             border-radius: 4px; cursor: pointer; margin: 10px 0; }
    .instructions { background-color: #e9f7ef; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
  </style>
</head>
<body>
  <h1>${projectName || 'Generated Apps Script'}</h1>
  
  <div class="instructions">
    <h2>How to use this code:</h2>
    <ol>
      <li>Click the "Download Code.gs" button below</li>
      <li>In Google Sheets, go to Extensions > Apps Script</li>
      <li>Replace the content in the Code.gs file with this code</li>
      <li>Save the project</li>
      <li>Run the function you want to execute</li>
    </ol>
  </div>
  
  <button class="button" onclick="downloadCode()">Download Code.gs</button>
  
  <h2>Code Preview:</h2>
  <pre id="codeContent">${code.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>
  
  <button class="button" onclick="downloadCode()">Download Code.gs</button>
</body>
</html>`;
    
    return {
      success: true,
      message: "Code generated successfully. Click the links to download the file.",
      code: code,
      htmlCode: htmlCode
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
