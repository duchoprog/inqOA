const fs = require("fs");
const fsPromises = fs.promises;
const path = require("path");
const { setTimeout } = require("timers");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");

// Function to save the file buffer to the uploads directory
async function saveFileToUploads(file) {
  const uploadsDir = path.join(__dirname, "uploads");

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, file.originalname);

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`File saved to ${filePath}`);
}

async function saveFileToFiles(file) {
  console.log("saving file");
  const uploadsDir = path.join(__dirname, "files");

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, file.originalname);

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`File saved to ${filePath}`);
}

async function writeOutputToExcel(responseArray, res) {
  // Process the data
  const processedData = await processData(responseArray);

  // Read the existing workbook
  const filePath = "./INQUIRY 2024 TEMPLATE v4 pablo.xlsx";
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Get the first sheet
  const sheetName = workbook.worksheets[0].name;
  const worksheet = workbook.getWorksheet(sheetName);

  // Define the starting row for the new data
  const startRow = 6;

  // Insert the new data starting from row 6
  processedData.forEach((rowData, index) => {
    const row = worksheet.getRow(startRow + index);
    Object.keys(rowData).forEach((key, colIndex) => {
      row.getCell(colIndex + 1).value = rowData[key];
    });
    row.commit();
  });

  // Preserve column widths
  const columnWidths = worksheet.columns.map((col) => col.width);
  worksheet.columns.forEach((col, index) => {
    col.width = columnWidths[index];
  });

  // Write the workbook back to the file
  await workbook.xlsx.writeFile("./output/products.xlsx");

  setTimeout(() => {}, 5000);

  await res.json({ success: true, redirectUrl: "/download" });
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

  responseArray.forEach((item) => {
    // Remove any content after the closing bracket '}]' but keep the closing single quote
    const regex = /【[^【】]*】/g;
    let cleanedResponse = item.openaiResponse.replace(regex, "");
    // Parse the JSON data
    let jsonData = JSON.parse(cleanedResponse);
    allData = allData.concat(jsonData);
  });

  return allData;
}

async function manageFolders(folders) {
  for (const folderName of folders) {
    const folderPath = path.resolve(__dirname, folderName);

    try {
      // Check if the folder exists
      await fsPromises.access(folderPath);
      // If it exists, delete it
      await fsPromises.rm(folderPath, { recursive: true });
      console.log(`Deleted folder: ${folderPath}`);
    } catch (error) {
      // If it doesn't exist, ignore the error
      if (error.code !== "ENOENT") {
        throw error;
      }
    }

    // Create the folder
    await fsPromises.mkdir(folderPath);
    console.log(`Created folder: ${folderPath}`);
  }
}

module.exports = {
  deleteAllFilesInDir,
  deleteFolder,
  saveFileToUploads,
  writeOutputToExcel,
  processData,
  manageFolders,
  saveFileToFiles,
};
