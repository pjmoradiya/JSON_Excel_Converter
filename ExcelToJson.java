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

/**
 * Scans an input folder for .xlsx files, converts them to JSON by:
 * 1) For each sheet: 
 *     - If there's exactly 1 data row (besides header), treat it as single object {"keyValue":[...]}.
 *     - If there are multiple data rows, treat it as array of objects [{"keyValue":[...]}, ...].
 * 2) Reassembles them under benefitRequest.dataSet.CVS with each sheet name as the key.
 * 3) Writes placeholders for transactionID, clientCode, data, etc.
 */
public class ExcelToJsonBatch {

    // Adjust as needed
    private static final String INPUT_DIR  = "C:/Development/myCode/inputExcel";
    private static final String OUTPUT_DIR = "C:/Development/myCode/outputJson";

    public static void main(String[] args) {
        try {
            // 1) Ensure input dir
            File inputDir = new File(INPUT_DIR);
            if (!inputDir.exists()) {
                boolean created = inputDir.mkdirs();
                if (created) {
                    System.out.println("Created input folder: " + INPUT_DIR);
                }
                System.out.println("Place your .xlsx files there, then run again.");
                return;
            }

            // 2) Find .xlsx
            File[] excelFiles = inputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));
            if (excelFiles == null || excelFiles.length == 0) {
                System.out.println("No .xlsx files found in " + INPUT_DIR);
                return;
            }

            // 3) Ensure output dir
            File outputDir = new File(OUTPUT_DIR);
            if (!outputDir.exists()) {
                boolean created = outputDir.mkdirs();
                if (created) {
                    System.out.println("Created output folder: " + OUTPUT_DIR);
                }
            }

            // 4) Convert each xlsx
            for (File excelFile : excelFiles) {
                String baseName = excelFile.getName().replaceFirst("[.][^.]+$", "");
                File jsonFile = new File(outputDir, baseName + ".json");

                System.out.println("Converting " + excelFile.getName() + " to " + jsonFile.getName());
                convertExcelToJson(excelFile, jsonFile);
            }

            System.out.println("All conversions finished.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Convert a single Excel file to JSON, scanning all sheets.
     */
    private static void convertExcelToJson(File excelFile, File jsonFile) {
        Workbook workbook = null;
        FileInputStream fis = null;
        FileOutputStream fos = null;

        try {
            fis = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(fis);

            // Build the structure:
            // {
            //   "benefitRequest": {
            //     "transactionID": "PLACEHOLDER",
            //     "clientCode": "PLACEHOLDER",
            //     "data": null,
            //     "dataSet": {
            //       "CVS": {
            //         [SHEET_NAME_1]: [OBJECT or ARRAY],
            //         [SHEET_NAME_2]: ...
            //       }
            //     }
            //   }
            // }

            JSONObject root = new JSONObject();
            JSONObject benefitRequest = new JSONObject();
            root.put("benefitRequest", benefitRequest);

            // placeholders for top-level
            benefitRequest.put("transactionID", "PLACEHOLDER");
            benefitRequest.put("clientCode", "PLACEHOLDER");
            benefitRequest.put("data", JSONObject.NULL);

            JSONObject dataSet = new JSONObject();
            JSONObject cvs = new JSONObject();
            dataSet.put("CVS", cvs);
            benefitRequest.put("dataSet", dataSet);

            // Parse each sheet
            int numSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < numSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                // Let's see how many data rows it has (besides header).
                int dataRowCount = sheet.getPhysicalNumberOfRows() - 1; 
                // Note: getPhysicalNumberOfRows() can be tricky if there are blank rows,
                // but let's assume we have a consistent spreadsheet.

                if (dataRowCount <= 0) {
                    // no data rows
                    continue;
                }
                else if (dataRowCount == 1) {
                    // single row -> interpret as single { "keyValue":[ ... ] }
                    JSONObject singleObject = new JSONObject();
                    parseKeyValueSingleObjectSheet(sheet, singleObject);
                    cvs.put(sheetName, singleObject);
                } else {
                    // multiple rows -> interpret as array [ { "keyValue":[...] }, ... ]
                    JSONArray arrayOfObjects = new JSONArray();
                    parseKeyValueArraySheet(sheet, arrayOfObjects);
                    cvs.put(sheetName, arrayOfObjects);
                }
            }

            // Write out JSON
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
     * Parses a sheet that has 1 header row + 1 data row.
     * Rebuilds an object: { "keyValue": [ { "attribute":..., "value":... }, ... ] }
     */
    private static void parseKeyValueSingleObjectSheet(Sheet sheet, JSONObject destination) {
        JSONArray keyValueArray = new JSONArray();

        Iterator<Row> rowIterator = sheet.iterator();
        if (!rowIterator.hasNext()) return; // no rows at all
        Row headerRow = rowIterator.next();

        if (!rowIterator.hasNext()) return; // no data row
        Row dataRow = rowIterator.next();

        int cellCount = headerRow.getPhysicalNumberOfCells();
        for (int c = 0; c < cellCount; c++) {
            Cell headerCell = headerRow.getCell(c);
            if (headerCell == null) continue;
            String attribute = headerCell.getStringCellValue();

            Cell dataCell = dataRow.getCell(c);
            String value = (dataCell == null) ? "" : dataCell.toString();

            JSONObject kv = new JSONObject();
            kv.put("attribute", attribute);
            kv.put("value", value);
            keyValueArray.put(kv);
        }

        destination.put("keyValue", keyValueArray);
    }

    /**
     * Parses a sheet that has 1 header row + multiple data rows.
     * Rebuilds an array: [ { "keyValue":[ ... ] }, { "keyValue":[ ... ] }, ... ]
     */
    private static void parseKeyValueArraySheet(Sheet sheet, JSONArray destinationArray) {
        Iterator<Row> rowIterator = sheet.iterator();
        if (!rowIterator.hasNext()) return;

        // header row
        Row headerRow = rowIterator.next();
        int cellCount = headerRow.getPhysicalNumberOfCells();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < cellCount; c++) {
            Cell cell = headerRow.getCell(c);
            if (cell == null) {
                headers.add("");
            } else {
                headers.add(cell.toString());
            }
        }

        // data rows
        while (rowIterator.hasNext()) {
            Row dataRow = rowIterator.next();
            JSONArray keyValueArray = new JSONArray();

            for (int c = 0; c < headers.size(); c++) {
                String attribute = headers.get(c);
                if (attribute == null || attribute.isEmpty()) continue;

                Cell dataCell = dataRow.getCell(c);
                String value = (dataCell == null) ? "" : dataCell.toString();

                JSONObject kv = new JSONObject();
                kv.put("attribute", attribute);
                kv.put("value", value);
                keyValueArray.put(kv);
            }

            JSONObject elementObj = new JSONObject();
            elementObj.put("keyValue", keyValueArray);
            destinationArray.put(elementObj);
        }
    }
}
