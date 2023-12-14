const express = require('express');
const router = express.Router();
const fileUpload = require('express-fileupload');
const multer = require('multer');
const path = require('path');
const ExcelJS = require("exceljs");

// Xử lý yêu cầu GET đến /random-sk-duyet
router.get('/', (req, res) => {
    // Thực hiện các xử lý khi nhận yêu cầu GET tới /random-sk-duyet
    res.render('other-content', { title: 'OTHER CONTENT' });
});

// Cấu hình Multer để lưu trữ tệp tin được tải lên vào thư mục 'uploads'
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/other-content');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

router.post('/upload', upload.single('uploadedFile'), async (req, res) => {
    if (!req.file) {
        return res.status(400).send('Không có tệp tin nào được tải lên.');
    }
    // Xử lý tệp tin nếu cần
    const uploadedFileName = `./uploads/other-content/${req.file.filename}`;
    let information = req.body

    const xlsxFilePath = uploadedFileName; // Tên file input
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(xlsxFilePath);
    // xử lý excel
    let sheet = workbook.getWorksheet("Sheet1");
    const rows = sheet.getRows(12, 500); // Lấy các hàng từ dòng 12 đến dòng 20
    let rowIndex = 0;
    console.log(information)
    for (const key in information) {
        if (Array.isArray(information[key])) {
            let date = information[key][0];
            content = information[key][1];
            amount = information[key][2];
            // console.log([date,content])
            
            // Duyệt qua từng hàng trong cột A và tìm kiếm mục cần thiết
            for (let index = 0; index < rows.length; index++) {
                let row = rows[index + 12];
                if(row == undefined)
                    continue;
                let cellValue = row.getCell('A').value;
                if (cellValue != undefined && cellValue===date) {
                    console.log(cellValue,content)
                    row.getCell('E').value = content
                    const  isDebit =  key.includes("feeAccount") ? true : false;
                    
                    if(isDebit == true) {
                        row.getCell('C').value=amount
                    }else{
                        row.getCell('D').value=amount
                    }

                   break;
                }
               
            }

        }
   
    }



    // Lưu file output
    const outputFilePath = `outputs/other-content/${req.file.filename}_other-content.xlsx`;
    await workbook.xlsx.writeFile(outputFilePath);
    res.json({
        message: 'success'
    })

})

module.exports = router;