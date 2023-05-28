const express = require("express");
const multer = require("multer");
const cors = require("cors");
const XLSX = require("xlsx");
const bodyParser = require("body-parser");
const { MongoClient } = require('mongodb');


const client = new MongoClient("mongodb+srv://sumbembaevb:P6TQio8zC978woo7@kvadder.sp0jven.mongodb.net/kvadder?retryWrites=true&w=majority")
let collectionsRepository = null;
const start = async () => {
    try {
        console.log('connection......')
        await client.connect();
        console.log('Соеденение с БД установлено!');
        collectionsRepository = client.db().collection("Collections");
        //await client.db().createCollection('collections');
    } catch (e) {
        console.log(e);
    }
}

start();
const app = express();
app.use(cors());
app.use("/files", express.static("uploads"));
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

const storage = multer.diskStorage({
  destination: function (req, file, callback) {
    callback(null, "./uploads");
  },
  filename: function (req, file, callback) {
    callback(null, file.originalname);
  },
});

const uploads = multer({ storage: storage }).single("file");

app.post("/uploads", async (req, res) => {
  uploads(req, res, function (err) {
    if (err) {
      console.log(err);
    } else {
      var FileName = req.file.filename;
      console.log(req.file);
      res.status(200).send(FileName);
    }
  });
});

app.get("/collections", async(req, res) => {
    const data = await collectionsRepository.find().toArray();
    res(data);        
})

app.post("/convert", async (req, res) => {
  //console.log(req.body.filename);
  const filename = req.body.filename;
  console.log(filename);
  console.log('ext(filename)', ext(filename));
  const extension = ext(filename);
  switch (extension) {
    case "csv":
      return res.status(200).send(filename);
      break;
    case "txt":
      return res.status(200).send(filename);
      break;
    case "xlsx":
      return res.status(200).send(convertFromXLSX(filename));
      break;
    case "xls":
      return res.status(200).send(convertFromXLSX(filename));
      break;

    default:
      break;
  }
});

app.listen(3000, function () {
  console.log("Start on 3000 PORT");
});

//utils
//Получить расширение файла
function ext(name) {
  return name.match(/\.([^.]+)$/)?.[1];
}

function convertFromXLSX(filename) {
  console.log('URAAAAAAAA', filename);
  const workBook = XLSX.readFile("./uploads/" + filename);
  const savedFileName = filename.replace(/\.[^.]+$/, ".csv");
  XLSX.writeFile(workBook, "./uploads/" + savedFileName, { bookType: "csv" });
  return savedFileName;
}
