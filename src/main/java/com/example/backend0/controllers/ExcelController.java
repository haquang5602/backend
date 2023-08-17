package com.example.backend0.controllers;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

@RestController
public class ExcelController {

  @GetMapping("/read-excel")
  public ResponseEntity<String> readExcel() {
    try {
      // Đường dẫn đến tệp Excel
      String excelFilePath = "Huong dan Projects.xlsx";

      // Mở tệp Excel bằng FileInputStream
      InputStream inputStream = new FileInputStream(excelFilePath);

      // Tạo Workbook từ tệp Excel
      Workbook workbook = new XSSFWorkbook(inputStream);

      // Chọn sheet cần đọc (ví dụ: sheet đầu tiên)
      Sheet sheet = workbook.getSheetAt(0);

      // Duyệt qua các dòng trong sheet
      for (Row row : sheet) {
        for (Cell cell : row) {
          // Xử lý từng ô (cell) tại đây
          String cellValue = cell.toString();
          System.out.print(cellValue + "\t");
        }
        System.out.println(); // Xuống dòng sau mỗi dòng
      }

      // Đóng Workbook và InputStream
      workbook.close();
      inputStream.close();

      return ResponseEntity.ok("Đã đọc dữ liệu từ tệp Excel thành công");
    } catch (IOException e) {
      e.printStackTrace();
      return ResponseEntity.status(500).body("Lỗi khi đọc tệp Excel");
    }
  }
}
