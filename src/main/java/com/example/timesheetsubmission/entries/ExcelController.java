package com.example.timesheetsubmission.entries;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.mail.MessagingException;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.Month;
import java.time.format.TextStyle;
import java.util.Locale;

@RestController
public class ExcelController {

    @Autowired
    private JavaMailSender emailSender;

    @PostMapping("/uploadExcel")
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            // Process the file
            Workbook workbook = WorkbookFactory.create(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0); // Assuming the timesheet is in the first sheet

            DataFormatter dataFormatter = new DataFormatter();
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            // Extract the required information
            String name = sheet.getRow(40).getCell(1).getStringCellValue(); // Name is in row 41, column B

            // Evaluate the formula in the total hours cell and get the result
            Cell totalHoursCell = sheet.getRow(35).getCell(4); // Total hours is in row 36, column E
            formulaEvaluator.evaluateFormulaCell(totalHoursCell);
            String totalHours = dataFormatter.formatCellValue(totalHoursCell);

            // Extract and format the period start date
            Cell periodCell = sheet.getRow(0).getCell(2); // Date is in row 1, column C
            formulaEvaluator.evaluateFormulaCell(periodCell);
            String period = dataFormatter.formatCellValue(periodCell);

            // Format the date correctly
            String[] temp = period.split("/");
            Month month = Month.of(Integer.parseInt(temp[0]));
            String monthName = month.getDisplayName(TextStyle.FULL, Locale.ENGLISH);
            String title = name + "-" + monthName + temp[2] + "-";

            // Create a new Excel file with the extracted information
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet(title + "TimeSheet");

            Row newRow = newSheet.createRow(0);
            newRow.createCell(0).setCellValue("Name:");
            newRow.createCell(1).setCellValue(name);

            newRow = newSheet.createRow(1);
            newRow.createCell(0).setCellValue("Total Hours:");
            newRow.createCell(1).setCellValue(totalHours);

            newRow = newSheet.createRow(2);
            newRow.createCell(0).setCellValue("Period Start:");
            newRow.createCell(1).setCellValue(period);

            // Write the new Excel file to a ByteArrayOutputStream
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            newWorkbook.write(bos);
            byte[] bytes = bos.toByteArray();

            // Email the new Excel file as an attachment
            sendEmailWithAttachment(title + "TimeSheet.xlsx", bytes);

            return ResponseEntity.ok("File uploaded successfully");
        } catch (IOException | MessagingException | EncryptedDocumentException e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to upload file");
        }
    }

    private void sendEmailWithAttachment(String fileName, byte[] fileBytes) throws MessagingException {
        jakarta.mail.internet.MimeMessage message = emailSender.createMimeMessage();
        MimeMessageHelper helper;
        try {
            helper = new MimeMessageHelper(message, true);
            helper.setFrom("tristanmcurtis844@gmail.com"); // Update with your email address
            helper.setTo("tmcurti4@ncsu.edu"); // Update with recipient's email address
            helper.setSubject("Timesheet Uploaded");
            helper.setText("Please find the attached timesheet file.");

            helper.addAttachment(fileName, new ByteArrayResource(fileBytes));
            emailSender.send(message);
        } catch (jakarta.mail.MessagingException e) {
            e.printStackTrace();
        }
    }

    @GetMapping("/success")
    public String successPage() {
        return "success";
    }
}
