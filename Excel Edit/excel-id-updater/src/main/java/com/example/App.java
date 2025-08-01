package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class App {
    public static void main(String[] args) throws Exception {
        Scanner scanner = new Scanner(System.in);
        System.out.print("Nhập STT bắt đầu cập nhật: ");
        int startStt = scanner.nextInt();

        String inputFile = "input.xlsx";   // Đặt file Excel gốc ở đây
        String outputFile = "output.xlsx"; // Kết quả sau xử lý

        FileInputStream fis = new FileInputStream(inputFile);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Lưu trữ thông tin các dòng hợp lệ
        List<RowData> validRows = new ArrayList<>();
        Map<Integer, Integer> oldToNewIdMap = new HashMap<>();
        
        int totalRows = sheet.getLastRowNum();
        
        // Bước 1: Thu thập tất cả dòng hợp lệ (không null và có dữ liệu)
        for (int i = 1; i <= totalRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(2) != null) { // Kiểm tra dòng và cột ID có tồn tại
                try {
                    int oldId = (int) row.getCell(2).getNumericCellValue();
                    int oldParentId = row.getCell(3) != null ? (int) row.getCell(3).getNumericCellValue() : 0;
                    
                    RowData rowData = new RowData();
                    rowData.row = row;
                    rowData.oldId = oldId;
                    rowData.oldParentId = oldParentId;
                    
                    validRows.add(rowData);
                } catch (Exception e) {
                    // Bỏ qua dòng có lỗi dữ liệu
                    System.out.println("⚠️ Bỏ qua dòng " + (i+1) + " do lỗi dữ liệu");
                }
            }
        }
        
        // Bước 2: Cập nhật STT và tạo mapping ID cho TẤT CẢ dòng
        int updatedCount = 0;
        for (int i = 0; i < validRows.size(); i++) {
            RowData rowData = validRows.get(i);
            int newStt = i + 1;  // STT mới
            
            // Cập nhật STT
            rowData.row.getCell(0).setCellValue(newStt);
            
            // Tạo mapping cho TẤT CẢ dòng
            if (newStt >= startStt) {
                // Cập nhật ID = STT cho dòng >= startStt
                int newId = newStt;  
                rowData.newId = newId;
                oldToNewIdMap.put(rowData.oldId, newId);
                updatedCount++;
            } else {
                // Giữ nguyên ID cũ cho dòng < startStt, nhưng vẫn map
                rowData.newId = rowData.oldId;
                oldToNewIdMap.put(rowData.oldId, rowData.oldId);
            }
        }
        
        // Bước 3: Cập nhật ID và ID_Parent
        for (RowData rowData : validRows) {
            Row row = rowData.row;
            
            // Cập nhật ID
            row.getCell(2).setCellValue(rowData.newId);
            
            // Cập nhật ID_Parent
            int newParentId = 0;
            if (rowData.oldParentId != 0) {
                if (oldToNewIdMap.containsKey(rowData.oldParentId)) {
                    newParentId = oldToNewIdMap.get(rowData.oldParentId);
                    System.out.println("🔄 Map Parent: " + rowData.oldParentId + " → " + newParentId + 
                                     " (cho ID cũ " + rowData.oldId + " → mới " + rowData.newId + ")");
                } else {
                    newParentId = rowData.oldParentId; // Giữ nguyên nếu không tìm thấy
                    System.out.println("⚠️ Giữ nguyên Parent ID " + rowData.oldParentId + 
                                     " (không tìm thấy mapping cho dòng ID " + rowData.oldId + ")");
                }
            }
            row.getCell(3).setCellValue(newParentId);
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(outputFile);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("✅ Đã cập nhật thành công!");
        System.out.println("📊 Tổng số dòng: " + validRows.size());
        System.out.println("🔄 Số dòng đã cập nhật ID: " + updatedCount + " (từ STT " + startStt + " trở xuống)");
        System.out.println("💾 File kết quả: " + outputFile);
    }
    
    // Class helper để lưu thông tin dòng
    static class RowData {
        Row row;
        int oldId;
        int oldParentId;
        int newId;
    }
}
