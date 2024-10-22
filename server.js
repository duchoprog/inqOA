const express = require("express");
const cors = require("cors");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
require("dotenv").config(); // Require dotenv configuration
const routes = require("./routes");
const ejs = require("ejs");

const OpenAI = require("openai");
const {
  deleteFolder,
  saveFileToUploads,
  writeOutputToExcel,
  manageFolders,
  saveFileToFiles,
  savePrevFileToExcelBase,
  deleteOneFile,
  createFolder,
} = require("./utilities.js");
const { processUploadedFile } = require("./assistantTest2.js");
const { getImages } = require("./extractImages.js");
const { replaceImages } = require("./replaceImages.js");
const { extractImageExcel } = require("./extractImageExcel.js");
const { extractImageDocx } = require("./extractImageDocx");
const {
  cleanText,
  deleteAllFilesInDir,
  deleteOldFolders,
} = require("./utilities.js");
const openai = new OpenAI();

const app = express();
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

let folderPath = "newproject";

app.use(cors());
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.json()); // For parsing application/json
app.use(express.urlencoded({ extended: true })); // For parsing application/x-www-form-urlencoded
app.use("/", routes());

app.use(express.static(path.join(__dirname, "public")));

// Serve frontend
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.post(
  "/submit",
  upload.fields([{ name: "files" }, { name: "previousFiles" }]),

  async (req, res) => {
    console.log("cu");
    deleteOldFolders("./");

    folderPath = await createFolder(req.projectName);
    const contentArray = req.body.content;
    const files = req.files["files"] || [];
    const previousFiles = req.files["previousFiles"] || [];
    const responsesArray = [];
    var d = new Date();
    d = d.getTime().toString();
    console.log("req.body:", req.body);
    req.body.sessionID = d;
    await manageFolders(req.body.sessionID);
    //console.log("req.body:", req.body);
    try {
      //await deleteOneFile("./excelBase/addInfoToThis.xlsx");

      if (previousFiles.length > 0) {
        let file = previousFiles[0];
        let fileContent = file.buffer;

        if (
          file.mimetype === "application/vnd.ms-excel" ||
          file.mimetype ===
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ) {
          await savePrevFileToExcelBase(file, req.body.sessionID);
        }
      }

      await processEachFile(req.body.sessionID);

      if (contentArray && contentArray.length > 0) {
        console.log("CONTENT");
        await processInputContent();
      }

      console.log("responsesArray", responsesArray);
      await writeOutputToExcel(
        responsesArray,
        res,
        req.body.projectName,
        req.body.sessionID
      );
      await replaceImages(req);
    } catch (error) {
      console.error("Error during /submit process:", error);
      res.json({
        success: true,
        redirectUrl: "/error",
        message: "An error occurred during the process.",
        sessionID: req.body.sessionID,
      });
      //res.status(500).send("An error occurred during the process.");
      return; // Stops further execution of the route
    }

    async function processEachFile(sessionID) {
      for (const file of files) {
        console.log(`comienza proceso de file ${file.originalname}`);
        //console.log("file:", file.originalname);
        let fileContent = file.buffer;
        let filePath = null;

        if (
          file.mimetype === "application/vnd.ms-excel" ||
          file.mimetype ===
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ) {
          try {
            await saveFileToFiles(file, sessionID);

            let excelPath = `./${req.body.sessionID}/files/${file.originalname}`;
            await extractImageExcel(excelPath, sessionID);

            const workbook = xlsx.read(file.buffer, { type: "buffer" });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            const trimmedData = data.slice(0, 100);
            const trimmedWorksheet = xlsx.utils.aoa_to_sheet(trimmedData);
            const fileContent = xlsx.utils.sheet_to_html(trimmedWorksheet);

            filePath = path.join(
              __dirname,
              req.body.sessionID,
              "uploads",
              `${file.originalname}.html`
            );
            fs.writeFileSync(filePath, fileContent);

            let openaiResponse = await processUploadedFile(
              filePath,
              req.body.resultsPerDoc,
              req.body.inquiry,
              res,
              req.body.sessionID
            );
            await handleImages(openaiResponse, req.body.sessionID);

            responsesArray.push(openaiResponse, req.body.sessionID);
          } catch (error) {
            console.error(`Error processing file ${file.originalname}:`, error);
            throw error; // Propagate error to stop execution
          }
        } else if (
          file.mimetype === "application/msword" ||
          file.mimetype ===
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ) {
          try {
            await saveFileToFiles(file, sessionID);
            await saveFileToUploads(file, req.body.sessionID);

            let docxPath = `./${req.body.sessionID}/files/${file.originalname}`;
            await extractImageDocx(docxPath, sessionID);
            ////
            let openaiResponse = await processUploadedFile(
              docxPath,
              req.body.resultsPerDoc,
              req.body.inquiry,
              res,
              req.body.sessionID
            );
            await handleImages(openaiResponse, req.body.sessionID);

            responsesArray.push(openaiResponse, req.body.sessionID);
          } catch (error) {
            console.error(`Error processing file ${file.originalname}:`, error);
            throw error; // Propagate error to stop execution
          }

          ////
        } else {
          try {
            await deleteAllFilesInDir(`./${sessionID}/images`);

            await saveFileToUploads(file, req.body.sessionID);
            filePath = `./${req.body.sessionID}/uploads/${file.originalname}`;

            if (filePath.includes(".pdf")) {
              await getImages(filePath, req);
            }

            let openaiResponse = await processUploadedFile(
              filePath,
              req.body.resultsPerDoc,
              req.body.inquiry,
              res,
              req.body.sessionID
            );

            console.log(openaiResponse);

            await handleImages(openaiResponse, req.body.sessionID);

            responsesArray.push(openaiResponse, req.body.sessionID);
          } catch (error) {
            console.error(
              `Error processing non-Excel file ${file.originalname}:`,
              error
            );
            throw error; // Propagate error to stop execution
          }
        }
      }
    }

    async function processInputContent() {
      for (let contentObj of contentArray) {
        try {
          const contentText = `Provider name: ${contentObj.title}\n${contentObj.content}`;
          const contentFilePath = path.join(
            __dirname,
            "uploads",
            `${contentObj.title}.txt`
          );
          fs.writeFileSync(contentFilePath, contentText);

          let openaiResponse = await processUploadedFile(
            contentFilePath,
            req.body.resultsPerDoc,
            req.body.inquiry,
            res
          );
          responsesArray.push(openaiResponse);
        } catch (error) {
          console.error("Error processing input content:", error);
          throw error; // Propagate error to stop execution
        }
      }
    }

    async function handleImages(openaiResponse) {
      openaiResponse = cleanText(openaiResponse);
      try {
        let imagesList = await fs.readdirSync(`./${req.body.sessionID}/images`);
        for (let i = 1; i <= imagesList.length; i++) {
          if (i === 1) {
            console.log(openaiResponse.openaiResponse);

            openaiResponse.openaiResponse =
              openaiResponse.openaiResponse.replace(
                `"PRODUCT REAL PICTURES": "NF"`,
                `"PRODUCT REAL PICTURES":"${imagesList[i - 1]}"`
              );
            openaiResponse.openaiResponse =
              openaiResponse.openaiResponse.replace(
                `"PRODUCT REAL PICTURES":"NF"`,
                `"PRODUCT REAL PICTURES":"${imagesList[i - 1]}"`
              );
          }
          openaiResponse.openaiResponse = openaiResponse.openaiResponse.replace(
            `"IMAGE ${i}": "NF"`,
            `"IMAGE ${i}":"${imagesList[i - 1]}"`
          );
          openaiResponse.openaiResponse = openaiResponse.openaiResponse.replace(
            `"IMAGE ${i}":"NF"`,
            `"IMAGE ${i}":"${imagesList[i - 1]}"`
          );
        }
        console.log("con imagenes: ", openaiResponse);
      } catch (error) {
        console.error("Error handling images:", error);
        throw error; // Propagate error to stop execution
      }
    }
  }
);

app.get("/download", (req, res) => {
  console.log("download is being hit");
  console.log(req.query);
  const resourceUrl = req.query.id; // Assuming the resource URL is passed as a query parameter
  const downPath = path.join(__dirname, resourceUrl, "output");
  res.render("download.ejs", { resourceUrl: resourceUrl });
});
app.get("/download2", (req, res) => {
  console.log("download is being hit");
  res.render("download2.ejs");
});

app.get("/error", (req, res) => {
  console.log("ERROR ROUTE");

  res.status(500).render("error.ejs");
});
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
