package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelToJsonBatch {

    // Adjust these paths as you prefer
    private static final String INPUT_DIR  = "C:/Development/myCode/inputExcel";
    private static final String OUTPUT_DIR = "C:/Development/myCode/outputJson";

    public static void main(String[] args) {
        try {
            // 1) Ensure input directory
            File inputDir = new File(INPUT_DIR);
            if (!inputDir.exists()) {
                boolean created = inputDir.mkdirs();
                if (created) {
                    System.out.println("Created input directory: " + INPUT_DIR);
                }
                System.out.println("Please place your .xlsx files in this directory, then run again.");
                return;
            }

            // 2) Find all .xlsx files
            File[] excelFiles = inputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));
            if (excelFiles == null || excelFiles.length == 0) {
                System.out.println("No .xlsx files found in " + INPUT_DIR);
                return;
            }

            // 3) Ensure output directory
            File outputDir = new File(OUTPUT_DIR);
            if (!outputDir.exists()) {
                boolean created = outputDir.mkdirs();
                if (created) {
                    System.out.println("Created output directory: " + OUTPUT_DIR);
                }
            }

            // 4) Convert each Excel file
            for (File excelFile : excelFiles) {
                String baseName = excelFile.getName().replaceFirst("[.][^.]+$", "");
                File jsonFile = new File(outputDir, baseName + ".json");

                System.out.println("Converting " + excelFile.getName() + " -> " + jsonFile.getName());
                convertExcelToJson(excelFile, jsonFile);
            }

            System.out.println("All Excel->JSON conversions finished.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Reads an Excel (.xlsx) file, looks for sheets "GeneralPlanDetails" & "Copay",
     * and reconstructs a JSON structure:
     * 
     * {
     *   "benefitRequest": {
     *     "transactionID": "PLACEHOLDER",
     *     "clientCode": "PLACEHOLDER",
     *     "data": null,
     *     "dataSet": {
     *       "CVS": {
     *         "GeneralPlanDetails": { "keyValue": [ ... ] },
     *         "Copay": [ { "keyValue": [ ... ] }, ... ]
     *       }
     *     }
     *   }
     * }
     */
    private static void convertExcelToJson(File excelFile, File jsonFile) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        Workbook workbook = null;

        try {
            fis = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(fis);

            // Build the main structure
            JSONObject root = new JSONObject();
            JSONObject benefitRequest = new JSONObject();
            root.put("benefitRequest", benefitRequest);

            // We do NOT have the original transactionID/clientCode/data.
            // We'll use placeholders or null. Adjust as needed.
            benefitRequest.put("transactionID", "PLACEHOLDER");
            benefitRequest.put("clientCode", "PLACEHOLDER");
            benefitRequest.put("data", JSONObject.NULL);

            // dataSet -> CVS
            JSONObject dataSet = new JSONObject();
            JSONObject cvs = new JSONObject();
            dataSet.put("CVS", cvs);
            benefitRequest.put("dataSet", dataSet);

            // Look for each sheet by name
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                if ("GeneralPlanDetails".equalsIgnoreCase(sheetName)) {
                    JSONObject generalPlanDetails = new JSONObject();
                    parseKeyValueObjectSheet(sheet, generalPlanDetails);
                    cvs.put("GeneralPlanDetails", generalPlanDetails);
                } 
                else if ("Copay".equalsIgnoreCase(sheetName)) {
                    JSONArray copayArray = new JSONArray();
                    parseKeyValueArraySheet(sheet, copayArray);
                    cvs.put("Copay", copayArray);
                }
            }

            // Write JSON to file
            fos = new FileOutputStream(jsonFile);
            fos.write(root.toString(4).getBytes());

            System.out.println("Created JSON: " + jsonFile.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try { workbook.close(); } catch (Exception e) { /* ignore */ }
            }
            if (fis != null) {
                try { fis.close(); } catch (Exception e) { /* ignore */ }
            }
            if (fos != null) {
                try { fos.close(); } catch (Exception e) { /* ignore */ }
            }
        }
    }

    /**
     * Parses a sheet that has exactly 1 header row + 1 data row.
     * Rebuilds: { "keyValue": [ { "attribute":"...", "value":"..." }, ... ] }
     */
    private static void parseKeyValueObjectSheet(Sheet sheet, JSONObject obj) {
        JSONArray keyValueArray = new JSONArray();

        Iterator<Row> rowIterator = sheet.iterator();
        if (!rowIterator.hasNext()) return;  // no rows
        Row headerRow = rowIterator.next();

        if (!rowIterator.hasNext()) return;  // no data
        Row dataRow = rowIterator.next();

        int lastCell = headerRow.getPhysicalNumberOfCells();
        for (int c = 0; c < lastCell; c++) {
            Cell headerCell = headerRow.getCell(c);
            if (headerCell == null) continue;
            String attribute = headerCell.getStringCellValue();

            Cell valueCell = dataRow.getCell(c);
            String value = (valueCell == null) ? "" : valueCell.toString();

            JSONObject kv = new JSONObject();
            kv.put("attribute", attribute);
            kv.put("value", value);
            keyValueArray.put(kv);
        }

        obj.put("keyValue", keyValueArray);
    }

    /**
     * Parses a sheet that has 1 header row + multiple data rows.
     * Rebuilds an array: [ { "keyValue": [ {...}, ... ] }, { ... }, ... ]
     */
    private static void parseKeyValueArraySheet(Sheet sheet, JSONArray array) {
        Iterator<Row> rowIterator = sheet.iterator();
        if (!rowIterator.hasNext()) return;

        // Header row
        Row headerRow = rowIterator.next();
        int lastCell = headerRow.getPhysicalNumberOfCells();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < lastCell; c++) {
            Cell cell = headerRow.getCell(c);
            headers.add(cell == null ? "" : cell.toString());
        }

        // Data rows
        while (rowIterator.hasNext()) {
            Row dataRow = rowIterator.next();
            JSONArray keyValueArray = new JSONArray();

            for (int c = 0; c < headers.size(); c++) {
                String attribute = headers.get(c);
                Cell cell = dataRow.getCell(c);
                String value = (cell == null) ? "" : cell.toString();

                JSONObject kv = new JSONObject();
                kv.put("attribute", attribute);
                kv.put("value", value);
                keyValueArray.put(kv);
            }

            JSONObject element = new JSONObject();
            element.put("keyValue", keyValueArray);
            array.put(element);
        }
    }
}
