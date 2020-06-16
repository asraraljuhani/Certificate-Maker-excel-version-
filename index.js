require('dotenv').config()
const express = require('express')
const app = express()
const path = require('path')
app.use(express.static(path.join(__dirname, 'public')))
var XLSX = require('xlsx')
var workbook = XLSX.readFile(`${__dirname}/test.xlsx`);
var sheet_name_list = workbook.SheetNames;
var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

const send = require('gmail-send')({
    user: process.env.EMAIL,
    pass: process.env.PASSWORD,
  });

xlData.forEach(e => {
    send({
            to:   e.email,
            subject: 'test 123',
            text:    `hi ${e.name}`, 
          })
          .then(({ result, full }) => console.log(result))
          .catch((error) => console.error('ERROR', error));      
})

app.listen(3000) 