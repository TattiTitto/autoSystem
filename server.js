const express = require('express')
const expressLayouts = require('express-ejs-layouts')
const dotenv = require('dotenv').config()
const { Document, Table, TableRow, TableCell, Paragraph, Packer, WidthType, ImageRun, TextRun } = require('docx');
const fs = require('fs')
const bodyParser = require('body-parser')
var HtmlDocx = require('html-docx-js');

var path = require('path')
var mammoth = require("mammoth")

const app = express()
const urlencodedParser = bodyParser.urlencoded({ extended: false })

app.set('view engine', 'ejs');
app.set('layout', './layouts/layout')
app.use(expressLayouts)
app.use(express.json())

app.get('/', (req, res) => {
    res.render('index', { title: "Excel | Docx"} )
})

app.post('/', urlencodedParser, (req, res) => {
    res.render('index', { title: "Excel | Docx"} )
    
    var text = fs.readFileSync("public/base/index.html", "utf8")
    text = text.split("--name--").join(req.body.item0)
    text = text.split("--obyekt--").join(req.body.item1)
    text = text.split("--address--").join(req.body.item2)
    text = text.split("--price--").join(req.body.item3)
    text = text.split("--vid--").join(req.body.item4)
    text = text.split("--from--").join(req.body.item5)
    text = text.split("--to--").join(req.body.item6)
    
    fs.writeFileSync("public/base/hisobot.html", text)

    var html = fs.readFileSync("public/base/hisobot.html", "utf8")
    
    var docx = HtmlDocx.asBlob(html);
    fs.writeFile('public/base/hisobot.docx',docx, function (err){
        if (err) return console.log(err);
        console.log('done');
    });
})

var options = {
    root: path.join(__dirname, 'public'),
    dotfiles: 'deny',
    headers: {
        'x-timestamp': Date.now(),
        'x-sent': true
    }
}
app.get('/save_docx', (req, res) => res.sendFile("base/hisobot.docx", options, function (err) {}))
// app.get('/save_excel', (req, res) => res.sendFile("excel.xlsx", options, function (err) {}))

app.listen(process.env.PORT, () => console.log(`${process.env.PORT} Listening port`))