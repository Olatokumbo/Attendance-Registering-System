const express=require("express");
const ejs = require("ejs");
var Excel = require('exceljs');
const arrayfirstName =[];
const arraylastName =[];
const arraygender =[];
const arrayparish =[];
const bodyParser = require("body-parser");
const app = express();
// A new Excel Work Book
var workbook = new Excel.Workbook();
// Some information about the Excel Work Book.
workbook.creator = 'David Odesola';
workbook.lastModifiedBy = '';
workbook.created = new Date(2018, 6, 19);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);
var i=0;
var sheet = workbook.addWorksheet('Sheet1');
    // A table header
    sheet.columns = [
        { header: 'Id', key: 'id' },
        { header: 'First Name', key: 'firstname1' },
        { header: 'Last Name', key: 'lastname1' },
        { header: 'Gender', key: 'gender1' },
        { header: 'Parish', key: 'parish1' }
    ]

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static(__dirname+"/public"));

app.get("/", function(req, res){
    res.render("index", {firstName: arrayfirstName, lastName: arraylastName,gender:arraygender, parish: arrayparish} );
});

app.post("/", function(req, res){
    i++;
    var firstname = req.body.firstName;
    var lastname = req.body.lastName;
    var gender = req.body.gender;
    var parish = req.body.parish;
    var date = new Date();
    console.log(firstname);
    console.log(lastname);
    console.log(gender);
    console.log(parish);
    console.log(date);
    sheet.addRow({id: i, firstname1: firstname, lastname1: lastname,gender1:gender, parish1: parish});
    // Save Excel on Hard Disk
    workbook.xlsx.writeFile("Attendance.xlsx");
//////
arrayfirstName.push(firstname);
arraylastName.push(lastname);
arraygender.push(gender);
arrayparish.push(parish);
res.redirect("/");
});
app.listen(8000, function(req, res){
    console.log("Listening at port 8000");
});