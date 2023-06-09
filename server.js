const express = require("express");
const multer = require("multer");
const cors = require("cors");
const XLSX = require("xlsx");
const bodyParser = require("body-parser");
const { MongoClient, ObjectId } = require("mongodb");

const client = new MongoClient(
  "mongodb+srv://sumbembaevb:P6TQio8zC978woo7@kvadder.sp0jven.mongodb.net/kvadder?retryWrites=true&w=majority"
);
let collectionsRepository = null;
let dataRepository = null;
const start = async () => {
  try {
    console.log("connection......");
    await client.connect();
    console.log("Соеденение с БД установлено!");
    collectionsRepository = client.db().collection("Collections");
    dataRepository = client.db().collection("CurrentData");
  } catch (e) {
    console.log(e);
  }
};

start();
const app = express();
app.use(cors({
  origin: '*',
  credentials: true,
}));
app.use("/files", express.static("uploads"));
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));

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
      res.status(200).send(FileName);
    }
  });
});

app.get("/collections", async (req, res) => {
  const data = await collectionsRepository.find().toArray();
  res.send(data);
});

app.get("/currentData", async (req, res) => {
  const result = await dataRepository.find().toArray();
  res.send(result);
});

app.post("/currentData", async(req, res) => {
  const result = await dataRepository.insertMany(req.body.tableData);
  res.send(result); 
});

app.post("/collections", async (req, res) => {
  const data = await collectionsRepository.insertOne(req.body);
  res.send(data);
});

app.put("/collections", async (req, res) => {
  console.log(req.body);
  const data = await collectionsRepository.updateOne(
    { label: req.body.label },
    { $set: {tags: req.body.tags} }
  );
  res.send(data);
});

app.delete("/collections", async (req, res) => {
  console.log('req.body', req.body);
  const data = await collectionsRepository.deleteOne(
    { _id: new ObjectId(req.body.ID) }
  );
  res.send(data);
});

app.post("/convert", async (req, res) => {
  const filename = req.body.filename;
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
  const workBook = XLSX.readFile("./uploads/" + filename);
  const savedFileName = filename.replace(/\.[^.]+$/, ".csv");
  XLSX.writeFile(workBook, "./uploads/" + savedFileName, { bookType: "csv" });
  return savedFileName;
}
