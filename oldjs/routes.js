const express = require("express");
const router = express.Router();
const controllers = require("./controllers");

const routes = () => {
  router.get("/download/excel", (req, res) => {
    const sessionID = req.query.resourceUrl;
    controllers.downloadExcel(req, res, sessionID);
  });
  return router;
};
module.exports = routes;
