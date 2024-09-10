const path = require("path");
const fs = require("fs");
const fsp = require("fs").promises;
const { getLastModifiedFile } = require("./lastFile.js");

const downloadExcel = async (req, res) => {
  console.log("Download Excel route is being hit");
  let dir = path.join(__dirname, req.query.resourceUrl, "output");
  let files = await fsp.readdir(dir);

  fileToDownload = path.join(
    __dirname,
    req.query.resourceUrl,
    "output",
    files[0]
  );
  console.log(files);
  console.log(fileToDownload);

  await fs.readFile(fileToDownload, (err, data) => {
    if (err) {
      res.status(500).send({ message: "Error al leer el archivo" });
    } else {
      console.log(fileToDownload);
      res.setHeader("Content-Disposition", `attachment; filename=${files[0]}`);
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.send(data);
    }
  });
};
module.exports = {
  downloadExcel,
};
