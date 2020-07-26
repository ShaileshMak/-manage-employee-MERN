const express = require("express");
const mongoose = require("mongoose");
const bodyParser = require("body-parser");

const config = require("./config");
const employees = require("./routes/api/employees");

const app = express();

app.use(bodyParser.json());

const db = config.mongoURI;

mongoose
  .connect(db, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
    useCreateIndex: true,
  })
  .then(() => console.log("DB connected!!"))
  .catch((error) => console.log(`Error connecting MongoDB!! ${error}`));

app.use("/api/employees/", employees);

const port = process.env.PORT || 5000;
app.listen(port, console.log(`App running on port ${port}`));
