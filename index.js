const express = require("express");
const multer = require('multer')
const path = require('path')
var sanitize = require("sanitize-filename");
const {exec} = require('child_process')
const fs = require('fs')

const app = express();

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/index.html")
});

const pdftodocxfilter = function (req, file, callback) {
  var ext = path.extname(file.originalname);
  if (ext !== ".pdf") {
    return callback("This Extension is not supported");
  }
  callback(null, true);
};

var pdftodocxstorage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "public/uploads");
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

var pdftodocxupload = multer({
  storage: pdftodocxstorage,
  fileFilter: pdftodocxfilter,
});

app.post("/pdftodocx", pdftodocxupload.single("file"), (req, res) => {
  if (req.file) {
    var pdf = req.file.path;
    console.log(pdf);

    var basename = path.basename(req.file.path);

    basename = sanitize(basename);

    console.log(basename);

    var htmlpath = `${basename.substr(0, basename.lastIndexOf("."))}.html`;

    htmlpath = sanitize(htmlpath);
    console.log(htmlpath);

    var docxpath = `${basename.substr(0, basename.lastIndexOf("."))}.docx`;

    docxpath = sanitize(docxpath);

    console.log(docxpath);

    exec(`soffice --convert-to html ${pdf}`, (err, stdout, stderr) => {
      if (err) {
        fs.unlinkSync(pdf);
        fs.unlinkSync(htmlpath);
        fs.unlinkSync(docxpath);
        res.send(err);
      }

      console.log(htmlpath);

      exec(
        `soffice --convert-to docx:'MS Word 2007 XML' ${htmlpath}`,
        (err, stdout, stderr) => {
          if (err) {
            fs.unlinkSync(pdf);
            fs.unlinkSync(htmlpath);
            fs.unlinkSync(docxpath);
            res.send(err);
          }
          console.log("output converted");
          console.log(docxpath);
          res.download(docxpath, (err) => {
            if (err) {
              fs.unlinkSync(pdf);
              fs.unlinkSync(htmlpath);
              fs.unlinkSync(docxpath);
              res.send(err);
            }

            fs.unlinkSync(pdf);
            fs.unlinkSync(htmlpath);
            fs.unlinkSync(docxpath);
          });
        }
      );
    });
  }
});

app.listen(5000,() => {
    console.log("App is listening on port 5000")
})
