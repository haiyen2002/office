const express = require('express');
const router = express.Router();
const fileUpload = require('express-fileupload');
const multer = require('multer');
const path = require('path');
const ExcelJS = require("exceljs");
const xlsx = require("xlsx");


router.get('/', (req, res) => {
    res.render('getName', { title: 'GET NAME' });
});
// Cấu hình Multer để lưu trữ tệp tin được tải lên vào thư mục 'uploads'
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/get-name');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

router.post('/upload',upload.single('uploadedFile'), async (req, res) => {
      if (!req.file) {
        return res.status(400).send('Không có tệp tin nào được tải lên.');
    }
    // Xử lý tệp tin nếu cần
    const uploadedFileName = `./uploads/get-name/${req.file.filename}`;
    let start = req.body.start;
    let end = req.body.end;
    let mainContent = req.body.mainContent;
    let splitContent = req.body.splitContent;
    let username = req.body.username;

    const xlsxFilePath = uploadedFileName; // Tên file input
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);
    // sheet gốc
    let worksheet = workbook.getWorksheet("Sheet1");
    
    // read w2
    const workbook2 = xlsx.readFile(xlsxFilePath);
    // Get the names of all sheets in the workbook.
    const sheetNames = workbook2.SheetNames;

    // sheet lấy data thứ tự bắt đầu từ 0
    const dataSheetName = sheetNames[2];

    // Get the first worksheet in the workbook.
    const worksheet2 = workbook2.Sheets[dataSheetName];
   // Convert the worksheet data into an array of objects.
   const data2 = xlsx.utils.sheet_to_json(worksheet2);
   const col1 = data2.map(item => item.name)
   const col2 = data2.map(item => item.column2).filter(item => item)
   const col3 = data2.map(item => item.column3).filter(item => item)

for(let i = start; i <= end; i++){
   let cellE = worksheet.getCell(`${mainContent}${i}`).text;
   worksheet.getCell(`${mainContent}${i}`).value = cellE;
   let cellG = worksheet.getCell(`${splitContent}${i}`);

   let result = ""
   
   if (cellE.includes("chuyen tien") && cellE.includes(";")){
       result = getStringBetweenOrToEnd(cellE, ";", "chuyen tien")
   }else if(cellE.includes("chuyen tien")){
       result = getStringBetweenOrToEnd(cellE, undefined, "chuyen tien")
   }else if(cellE.includes("chuyen khoan")){
       result = getStringBetweenOrToEnd(cellE, "-", "chuyen khoan")
   }else if(cellE.includes("OP VINFAST1-B2C PAYMENT")){
    result ="VU TRUC TUYEN ONEPAY".trim()
   }
   else{
       result = ""
   }

   cellG.value = result.includes(`${username}`) ? result.replace(`${username}`, col1[parseInt(Math.random() * col1.length)]).trim() : result.trim()
}
        // Lưu file output
    const outputFilePath = `outputs/get-name/${req.file.filename}_get-name.xlsx`;
    await workbook.xlsx.writeFile(outputFilePath);
        res.json({
            message: 'success'
        })
     
})


function getStringBetweenOrToEnd(input, startString, endString) {
    var startIndex = input.indexOf(startString);
    
    if (startIndex === -1 || !startString) {
      var endIndex = input.indexOf(endString);
      if (endIndex === -1) {
        return "End string not found in the input";
      }
      return input.substring(0, endIndex);
    }
  
    startIndex += startString.length;
  
    var endIndex = input.indexOf(endString, startIndex);
    if (endIndex === -1) {
      return "End string not found in the input after start string";
    }
  
    return Math.random() * 10 > 6 ? "VND-TGTT-" + input.substring(startIndex, endIndex) + "VIETNAM" : input.substring(startIndex, endIndex);
  }
module.exports = router;