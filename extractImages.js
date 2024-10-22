require("dotenv").config(); // Require dotenv configuration

const convertAPI = require("convertapi")(process.env.CONVERT_API_KEY);
const { manageFolders, deleteAllFilesInDir } = require("./utilities");

// Convert the PDF to images
async function getImages(filename, req) {
  console.log("extracting images from", filename);
  await deleteAllFilesInDir(`./${req.body.sessionID}/images`);
  filename = filename.trim().replaceAll("\\", "/");
  pdfFilePath = filename;
  console.log("pdffilepat: ", pdfFilePath);

  //todo: descomentar esto, lo comente para hacer pruebas
  try {
    const result = await convertAPI.convert(
      "extract-images",
      {
        File: pdfFilePath,
        ImageFormat: "jpg",
      },
      "pdf"
    );
    await result.saveFiles(`./${req.body.sessionID}/images`);
    await result.saveFiles(`./${req.body.sessionID}/imageVault`);
  } catch (error) {
    console.error("Error extracting images:", error);
  }
}

module.exports = {
  getImages,
};
