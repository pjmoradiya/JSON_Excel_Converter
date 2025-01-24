package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.json.JSONArray;
import org.json.JSONObject;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

/**
 * Scans an input folder for JSON files, creates one Excel file per JSON.
 * The JSON structure is assumed to be:
 * {
 *   "benefitRequest": {
 *     "transactionID": ...,
 *     "clientCode": ...,
 *     "data": ...,
 *     "dataSet": {
 *       "CVS": {
 *         "object1": { "keyValue": [ ... ] },
 *         "object2": [ { "keyValue": [...] }, { "keyValue": [...] } ],
 *         ...
 *       }
 *     }
 *   }
 * }
 * 
 * We skip a dedicated sheet for "benefitRequest" and only create sheets for each child object under CVS.
 * If the child is a single object with "keyValue", we create a single-row sheet.
 * If the child is an array, we create a multi-row sheet.
 */
public class JsonToExcelBatch {

    // Adjust input/output directories as needed
    private static final String INPUT_DIR  = "C:/Development/myCode/inputJson";
    private static final String OUTPUT_DIR = "C:/Development/myCode/outputExcel";

    public static void main(String[] args) {
        try {
            // 1) Ensure input directory exists
            File inputDir = new File(INPUT_DIR);
            if (!inputDir.exists()) {
                boolean created = inputDir.mkdirs();
                if (created) {
                    System.out.println("Created input directory: " + INPUT_DIR);
                }
                System.out.println("Please place your .json files in that folder, then run again.");
                return;
            }

            // 2) Gather all .json files in the input directory
            File[] jsonFiles = inputDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".json"));
            if (jsonFiles == null || jsonFiles.length == 0) {
                System.out.println("No JSON files found in " + INPUT_DIR);
                return;
            }

            // 3) Ensure output directory exists
            File outputDir = new File(OUTPUT_DIR);
            if (!outputDir.exists()) {
                boolean created = outputDir.mkdirs();
                if (created) {
                    System.out.println("Created output directory: " + OUTPUT_DIR);
                }
            }

            // 4) Convert each JSON file in a loop
            for (File jsonFile : jsonFiles) {
                String baseName = jsonFile.getName().replaceFirst("[.][^.]+$", "");
                File excelFile = new File(outputDir, baseName + ".xlsx");

                System.out.println("Converting " + jsonFile.getName() + " to " + excelFile.getName());
                convertJsonToExcel(jsonFile, excelFile);
            }

            System.out.println("All conversions finished.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Converts a single JSON file to Excel.
     */
    private static void convertJsonToExcel(File jsonFile, File excelFile) {
        Workbook workbook = null;
        FileOutputStream fos = null;

        try {
            // Read JSON from file
            String jsonString = new String(Files.readAllBytes(Paths.get(jsonFile.toURI())));
            JSONObject root = new JSONObject(jsonString);

            // Dig into the structure: benefitRequest -> dataSet -> CVS
            JSONObject benefitRequest = root.optJSONObject("benefitRequest");
            if (benefitRequest == null) {
                System.out.println("No 'benefitRequest' object found in " + jsonFile.getName());
                return;
            }
            JSONObject dataSet = benefitRequest.optJSONObject("dataSet");
            if (dataSet == null) {
                System.out.println("No 'dataSet' object found in " + jsonFile.getName());
                return;
            }
            JSONObject cvs = dataSet.optJSONObject("CVS");
            if (cvs == null) {
                System.out.println("No 'CVS' object found in " + jsonFile.getName());
                return;
            }

            // Create a workbook
            workbook = new XSSFWorkbook();

            // We'll create a custom CellStyle for the header row (tan color + borders).
            CellStyle headerStyle = createHeaderStyle(workbook);

            // For each key under CVS, check if it is an object or array, then create a sheet.
            for (String childKey : cvs.keySet()) {
                Object childValue = cvs.get(childKey);

                // If it's a single object with "keyValue", create a single-row sheet
                if (childValue instanceof JSONObject) {
                    JSONObject obj = (JSONObject) childValue;
                    if (obj.has("keyValue")) {
                        createSheetFromKeyValueObject(workbook, childKey, obj, headerStyle);
                    } else {
                        // Not our expected structure, skip or handle differently
                        System.out.println("Skipping " + childKey + ": no 'keyValue' found.");
                    }
                }
                // If it's an array, we assume each element has "keyValue"
                else if (childValue instanceof JSONArray) {
                    JSONArray arr = (JSONArray) childValue;
                    createSheetFromKeyValueArray(workbook, childKey, arr, headerStyle);
                }
                else {
                    System.out.println("Skipping " + childKey + ": unrecognized type (not object/array).");
                }
            }

            // Write the workbook
            fos = new FileOutputStream(excelFile);
            workbook.write(fos);
            System.out.println("Created Excel: " + excelFile.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Cleanup
            if (workbook != null) {
                try { workbook.close(); } catch (Exception e) { /* ignore */ }
            }
            if (fos != null) {
                try { fos.close(); } catch (Exception e) { /* ignore */ }
            }
        }
    }

    /**
     * Create a sheet from a single object with "keyValue": [ {attribute:..., value:...}, ... ]
     * We'll make exactly one data row.
     */
    private static void createSheetFromKeyValueObject(Workbook workbook, String sheetName,
                                                      JSONObject object, CellStyle headerStyle) {
        JSONArray keyValue = object.optJSONArray("keyValue");
        if (keyValue == null) return;

        // Collect attributes
        List<String> attributes = new ArrayList<>();
        for (int i = 0; i < keyValue.length(); i++) {
            JSONObject kv = keyValue.getJSONObject(i);
            attributes.add(kv.getString("attribute"));
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // HEADER ROW
        Row headerRow = sheet.createRow(0);
        for (int c = 0; c < attributes.size(); c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(attributes.get(c));
            cell.setCellStyle(headerStyle);
        }

        // DATA ROW
        Row dataRow = sheet.createRow(1);
        for (int i = 0; i < keyValue.length(); i++) {
            JSONObject kv = keyValue.getJSONObject(i);
            String attr  = kv.getString("attribute");
            String value = kv.optString("value", "");
            int colIndex = attributes.indexOf(attr);
            if (colIndex >= 0) {
                Cell cell = dataRow.createCell(colIndex);
                cell.setCellValue(value);
            }
        }

        // Auto-size
        for (int c = 0; c < attributes.size(); c++) {
            sheet.autoSizeColumn(c);
        }
    }

    /**
     * Create a sheet from an array of objects, each having a "keyValue" array.
     * Each element -> one row in the sheet.
     */
    private static void createSheetFromKeyValueArray(Workbook workbook, String sheetName,
                                                     JSONArray array, CellStyle headerStyle) {
        // Step 1: Collect all possible attributes across all array elements
        Set<String> allAttributes = new LinkedHashSet<>();
        List<JSONArray> rowsData = new ArrayList<>();

        for (int i = 0; i < array.length(); i++) {
            JSONObject element = array.optJSONObject(i);
            if (element == null) continue;
            JSONArray kvArr = element.optJSONArray("keyValue");
            if (kvArr != null) {
                rowsData.add(kvArr);
                for (int j = 0; j < kvArr.length(); j++) {
                    JSONObject kv = kvArr.getJSONObject(j);
                    allAttributes.add(kv.getString("attribute"));
                }
            }
        }

        if (allAttributes.isEmpty()) {
            // No data, skip creating sheet
            return;
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // Step 2: Header row
        List<String> attributeList = new ArrayList<>(allAttributes);
        Row headerRow = sheet.createRow(0);
        for (int c = 0; c < attributeList.size(); c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(attributeList.get(c));
            cell.setCellStyle(headerStyle);
        }

        // Step 3: One row per element
        int rowIndex = 1;
        for (JSONArray kvArr : rowsData) {
            Row row = sheet.createRow(rowIndex++);
            for (int j = 0; j < kvArr.length(); j++) {
                JSONObject kv = kvArr.getJSONObject(j);
                String attr  = kv.getString("attribute");
                String value = kv.optString("value", "");
                int colIndex = attributeList.indexOf(attr);
                if (colIndex >= 0) {
                    row.createCell(colIndex).setCellValue(value);
                }
            }
        }

        // Auto-size
        for (int c = 0; c < attributeList.size(); c++) {
            sheet.autoSizeColumn(c);
        }
    }

    /**
     * Creates a tan-colored header style with thin borders.
     */
    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Fill color: TAN
        style.setFillForegroundColor(IndexedColors.TAN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Thin borders
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // Optional: create a bold font
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }
}
