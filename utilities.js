const fs = require("fs");
const fsPromises = fs.promises;
const path = require("path");
const { setTimeout } = require("timers");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");

let folderPath = "newproject";

// Function to save the file buffer to the uploads directory
async function saveFileToUploads(file, sessionID) {
  const uploadsDir = path.join(__dirname, sessionID, "uploads");

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, file.originalname);

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`1 File saved to ${filePath}`);
}

async function saveFileToFiles(file, sessionID) {
  console.log("saving file");
  const uploadsDir = path.join(__dirname, sessionID, "files");

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, file.originalname);

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`2 File saved to ${filePath}`);
}

async function savePrevFileToExcelBase(file, sessionID) {
  console.log("saving file");
  const files = await fsPromises.readdir(path.join(__dirname, sessionID));
  console.log("files", files);

  const uploadsDir = path.join(__dirname, sessionID, `excelBase`);
  //const uploadsDir = folderPath + `/${sessionID}/excelBase`;
  console.log(uploadsDir);

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, "addInfoToThis.xlsx");

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`3 File saved to ${filePath}`);
}
async function writeOutputToExcel(responseArray, res, projectName, sessionID) {
  // Process the data
  const processedData = await processData(responseArray);

  // Select starting workbook
  //const filePath = "./INQUIRY 2024 TEMPLATE v4 pablo.xlsx";
  //let startingFiles = await fs.readdirSync("./excelBase");
  //console.log("starting files:", startingFiles.length);

  const filePath = fs.existsSync(`./${sessionID}/excelBase/addInfoToThis.xlsx`)
    ? `./${sessionID}/excelBase/addInfoToThis.xlsx`
    : "./excelBase/INQUIRY 2024 TEMPLATE v4 pablo2.xlsx";

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Get the first sheet
  const sheetName = workbook.worksheets[0].name;
  const worksheet = workbook.getWorksheet(sheetName);
  //add project id
  worksheet.getCell("H1").value = projectName;

  // Define the starting row for the new data */
  let startRow = 1;
  let isRowEmpty = false;

  while (!isRowEmpty) {
    isRowEmpty = true;
    const row = worksheet.getRow(startRow);

    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = row.getCell(col);
      if (cell.value !== null && !cell.formula) {
        isRowEmpty = false;
        break;
      }
    }

    if (!isRowEmpty) {
      startRow++;
    }
  }
  ////////////
  let columnsToFill = [
    3, 4, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
    25, 26, 27, 41, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55,
  ];

  processedData.forEach((rowData, index) => {
    console.log("start row en for each util108", startRow);

    const row = worksheet.getRow(startRow + index);
    let dataIndex = 0;

    columnsToFill.forEach((colIndex) => {
      if (dataIndex < Object.keys(rowData).length) {
        const cell = row.getCell(colIndex);
        // Check if the cell contains a formula
        if (!cell.formula) {
          cell.value = rowData[Object.keys(rowData)[dataIndex]];
          dataIndex++;
        } else {
          // Handle shared formulas by copying the formula from the master cell
          const masterCell = worksheet.getCell(startRow + index - 1, colIndex);
          if (masterCell.formula) {
            cell.formula = masterCell.formula;
          }
        }
      }
    });

    row.commit();
  });

  // Preserve column widths
  const columnWidths = worksheet.columns.map((col) => col.width);
  worksheet.columns.forEach((col, index) => {
    col.width = columnWidths[index];
  });

  //hide columns

  let hiddenCols = [
    "AC",
    "AD",
    "AE",
    "AF",
    "AG",
    "AJ",
    "AK",
    "AL",
    "AM",
    "AN",
    "AP",
    "AQ",
    "AR",
    "AS",
  ];

  for (col of hiddenCols) {
    let colToHide = worksheet.getColumn(col);
    colToHide.hidden = true;
  }

  // Write the workbook back to the file
  var d = new Date();
  d = d.getTime().toString();
  await workbook.xlsx.writeFile(
    `./${sessionID}/output/${projectName}${d}.xlsx`
  );

  setTimeout(() => {}, 5000);
  console.log("culo utilities");

  await res.json({ success: true, redirectUrl: `/download?id=${sessionID}` });
}

//DELETE ALL FILES IN FOLDER
async function deleteAllFilesInDir(dirPath) {
  try {
    console.log("deleting ", dirPath);
    fs.readdirSync(dirPath).forEach((file) => {
      console.log("deleting", `${dirPath}${file}`);
      fs.rmSync(path.join(dirPath, file));
    });
  } catch (error) {
    console.log(error);
  }
}
async function deleteOneFile(file) {
  if (fs.existsSync(file)) {
    try {
      fs.unlinkSync(file);
      console.log("File deleted successfully");
    } catch (error) {
      console.error("Error deleting the file:", error);
    }
  } else {
    console.log("File does not exist.");
  }
}
//CREATE FOLDER
async function createFolder(name) {
  var d = new Date();
  d = d.getTime().toString();
  folderPath = path.join(__dirname, `projects/${d}${name}`);

  try {
    fs.mkdirSync(folderPath);
    console.log("Folder ", folderPath, " created successfully!");
    return folderPath;
  } catch (err) {
    console.error("Error creating folder:", err);
  }
}
//DELETE FOLDER
async function deleteFolder(folder) {
  const directoryPath = path.resolve(__dirname, folder);

  // Check if the directory exists
  if (fs.existsSync(directoryPath)) {
    // Delete the directory if it exists
    try {
      fs.rmSync(directoryPath, { recursive: true });
      console.log(`${directoryPath} is deleted!`);
    } catch (err) {
      console.error(`Error while deleting ${directoryPath}.`, err);
    }
  } else {
    console.log(`${directoryPath} does not exist.`);
  }
}

// Function to process the data
async function processData(responseArray) {
  let allData = [];
  console.log("responsearray,", responseArray);

  responseArray.forEach((item) => {
    // Remove any content after the closing bracket '}]' but keep the closing single quote
    const regex = /【[^【】]*】/g;
    if (item.openaiResponse) {
      let cleanedResponse = item.openaiResponse.replace(regex, "");
      // Parse the JSON data
      let jsonData = JSON.parse(cleanedResponse);
      allData = allData.concat(jsonData);
    }
  });

  return allData;
}

async function manageFolders(sessionID) {
  const folders = [
    "images",
    "uploads",
    "output",
    "imageVault",
    "files",
    "excelBase",
  ];
  console.log("sessionID", sessionID);
  const folderPath = path.join(__dirname, sessionID);
  console.log(folderPath);

  await fsPromises.mkdir(folderPath);

  for (const folderName of folders) {
    const folderPath = path.resolve(__dirname, sessionID, folderName);
    console.log(folderPath);

    // Create the folder
    await fsPromises.mkdir(folderPath);
    console.log(`Created folder: ${folderPath}`);
  }
}

function cleanText(dirtyText) {
  // Step 1: Extract substrings between a colon and a comma or a closing curly bracket
  const regex = /:\s*([^,}]*)[,\}]/g;
  let match;
  let text;
  while ((match = regex.exec(dirtyText)) !== null) {
    if ((match[1].match(/"/g) || []).length > 2) {
      console.log("mal! ", match[1]);
      const split = match[1].split('"');
      let singleQuote = `${split.join("'")}`;
      singleQuote = `"${singleQuote.slice(1, -1)}"`;
      console.log(singleQuote);
      text = dirtyText.replace(match[1], singleQuote);
    } else {
      text = dirtyText;
    }
  }
  return text;
}

module.exports = {
  deleteAllFilesInDir,
  deleteFolder,
  saveFileToUploads,
  writeOutputToExcel,
  processData,
  manageFolders,
  saveFileToFiles,
  savePrevFileToExcelBase,
  deleteOneFile,
  createFolder,
  cleanText,
};
