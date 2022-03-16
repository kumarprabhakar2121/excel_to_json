const express = require("express");
const app = express();
const multer = require("multer");
const reader = require("xlsx");
const path = require("path");
app.use(express.json());

// store excel files into one folder
var storage = multer.diskStorage({
  destination: "./uploads/",
  filename: function (req, file, cb) {
    cb(
      null,
      file.originalname + "_" + Date.now() + path.extname(file.originalname)
    );
  },
});
var upload = multer({
  storage: storage,
}).single("excel");

app.post("/upload", async (req, res) => {
  var fileLocation;
  upload(req, res, (err) => {
    if (err) return res.json(err);
    // console.log(req.file);
    fileLocation = req.file.path;
    const file = reader.readFile(fileLocation);
    // console.log(file);
    let data = [];
    var emails = [];
    const sheets = file.SheetNames;

    for (let i = 0; i < sheets.length; i++) {
      const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
      temp.forEach((res) => {
        data.push(res);
      });
    }

    for (let i = 0; i < data.length; i++) {
      var tempEmail = data[i].Email;
      emails.push(tempEmail);
    }
    // Printing data from excel
    res.send(data);
    
    // Printing email field only from the excel
    // console.log(emails);
    // res.send(emails);
  });
});

app.listen(4000, () => {
  console.log("server listening on port 4000");
});
