const express = require('express');
const router = express.Router();
const fileUpload = require('express-fileupload');
const multer = require('multer');
const path = require('path');
const ExcelJS = require("exceljs");
const xlsx = require("xlsx");

router.get('/', (req, res) => {
  res.render('paymentRegular', { title: 'paymentRegular' });
});

// Cấu hình Multer để lưu trữ tệp tin được tải lên vào thư mục 'uploads'
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/paymentRegular');
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  }
});

const upload = multer({ storage: storage });

// *********XỬ LÝ FILE
router.post('/upload', upload.single('uploadedFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).send('Không có tệp tin nào được tải lên.');
  }


  // Xử lý tệp tin nếu cần
  const uploadedFileName = `./uploads/paymentRegular/${req.file.filename}`;
  const xlsxFilePath = uploadedFileName;
  const data = req.body;
  let workbook = new ExcelJS.Workbook();
  let message = "Success";
  let outputFilePath = "";
  try {
    // Tên file input

    await workbook.xlsx.readFile(xlsxFilePath);
    // sheet gốc
    let sheet = workbook.getWorksheet("Sheet1");

    const rows = sheet.getRows(data.start, data.end); // Lấy các hàng từ dòng 12 đến dòng 20
    console.log(data)

    // console.log(,)
    // Chuyển đổi chuỗi ngày thành đối tượng Date
    let startDay = data.startDate.split('/').reverse().join('-');
    let endDay = data.endDate.split('/').reverse().join('-');

    let currentDate = new Date(startDay);
    let finalDate = new Date(endDay);

    // Function để tăng ngày lên 1
    function addDays(date, days) {
      let result = new Date(date);
      result.setDate(result.getDate() + days);
      return result;
    }

    function toDateString(date) {
      // Lấy thông tin ngày, tháng, năm
      let day = currentDate.getDate();
      let month = currentDate.getMonth() + 1; // Tháng trong JavaScript đếm từ 0
      let year = currentDate.getFullYear();

      // Đảm bảo rằng chuỗi có đủ số 0 đằng trước nếu cần thiết
      day = day < 10 ? '0' + day : day;
      month = month < 10 ? '0' + month : month;

      // Tạo chuỗi theo định dạng 'dd/mm/yyyy'
      return `${day}/${month}/${year}`;
    }
    let countRead = 0;
    // Duyệt qua từng ngày và làm việc với từng ngày
    while (currentDate <= finalDate) {
      console.log(currentDate.toDateString()); // In ra màn hình ngày hiện tại
      countRead = countRead + 1;

      let dateStr = toDateString(currentDate);
      console.log("read row :", countRead, dateStr)
      for (let row of rows) {
        // console.log(row)
        if (row != undefined && row.getCell(data.columnDate).value.includes(dateStr)) {
          row.getCell(data.columnAmout).value = data.amount
          row.getCell(data.columnContent).value = data.content
          row.getCell(data.columnDebit).value = ""

          break;
        }
      }

      // Tăng ngày lên 1 cho lần lặp tiếp theo
      currentDate = addDays(currentDate, 1);
    }



    // Lưu file output
    outputFilePath = `outputs/paymentRegular/${req.file.filename}`;
    await workbook.xlsx.writeFile(outputFilePath);


  } catch (error) {
    console.log(error)
    if (error.code == 'EBUSY') {
      message = "File chưa đóng vk ơi"
    } else {
      message = "Lỗi chi rk biết:" + error
    }
  } finally {
    workbook = null
  }
  res.json({
    message,
    outputFilePath,
  })
});


module.exports = router;
