// Requiring the module
const express = require("express");
const reader = require("xlsx");
const formidable = require("formidable");
const bodyParser = require("body-parser");

const app = express();

app.set("port", process.env.PORT || 3000);

app.use(express.json());
app.use(express.urlencoded());

app.use(
  bodyParser.urlencoded({
    extended: false,
  })
);

app.use(bodyParser.json());

app.use(express.static("assets"));
app.use(express.static("bootstrap"));
app.use(express.static("public"));

// [*] Configuring Routes.
app.get("/", (req, res) => {
  // res.sendFile(__dirname + "/public/index.html");
  res.end("Home route");
});

app.get("/test", (req, res) => {
  res.end("This is the test route and the xlsx server is up and running");
});

app.use("/uploadsheets/", (req, res, next) => {
  res.header("Access-Control-Allow-OrIgin", "*");
  next();
});

app.post("/uploadsheets/", (req, res) => {
  const form = new formidable.IncomingForm();
  form.parse(req, function (err, fields, files) {
    const oldPath = files.sheet.filepath;

    let data = sheets(oldPath);

    res.send(JSON.stringify(data));
  });
});

function sheets(_file) {
  // Reading our test file
  const file = reader.readFile(_file);

  let data = [];

  const sheets = file.SheetNames;

  for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      data.push(res);
    });
  }

  return data;
}

app.listen(app.get("port"), () =>
  console.log("xlsx server up on port: ", app.get("port"))
);
