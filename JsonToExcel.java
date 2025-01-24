package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

public class JsonToExcelBatch {

    // Adjust these paths as you prefer
    private static final String INPUT_DIR  = "C:/Development/myCode/inputJson";
    private static final String OUTPUT_DIR = "C:/Development/myCode/outputExcel";

    public static void main(String[] args) {
        try {
            // 1) Ensure input directory exists; if not, create it and prompt user.
            File inputDir = new File(INPUT_DIR);
            if (!inputDir.exists()) {
                boolean created = inputDir.mkdirs();
                if (created) {
                    System.out.println("Created input directory: " + INPUT_DIR);
                }
                System.out.println("Please place your .json files in this directory, then run again.");
                return; 
            }

            // 2) Find all .json files
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

            // 4) Convert each JSON file
            for (File jsonFile : jsonFiles) {
                // Derive output .xlsx filename
                String baseName = jsonFile.getName().replaceFirst("[.][^.]+$", "");
                File excelFile = new File(outputDir, baseName + ".xlsx");

                System.out.println("Converting " + jsonFile.getName() + " -> " + excelFile.getName());
                convertJsonToExcel(jsonFile, excelFile);
            }

            System.out.println("All conversions finished.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Reads a single JSON file, and creates an Excel file
     * with sheets for "GeneralPlanDetails" and "Copay" only.
     * 
     * We skip creating a "benefitRequest" sheet, but you can modify
     * if you want to store it in a hidden sheet or somewhere else.
     */
    private static void convertJsonToExcel(File jsonFile, File excelFile) {
        Workbook workbook = null;
        FileOutputStream fos = null;

        try {
            // Read JSON
            String jsonString = new String(Files.readAllBytes(Paths.get(jsonFile.toURI())));
            JSONObject root = new JSONObject(jsonString);

            // "benefitRequest"
            JSONObject benefitRequest = root.optJSONObject("benefitRequest");
            if (benefitRequest == null) {
                // If there's no benefitRequest, just skip or handle differently
                System.out.println("No 'benefitRequest' object found in " + jsonFile.getName());
                return;
            }

            // dataSet -> CVS
            JSONObject dataSet = benefitRequest.optJSONObject("dataSet");
            if (dataSet == null) {
                System.out.println("No 'dataSet' in " + jsonFile.getName());
                return;
            }

            JSONObject cvs = dataSet.optJSONObject("CVS");
            if (cvs == null) {
                System.out.println("No 'CVS' in " + jsonFile.getName());
                return;
            }

            // Create workbook
            workbook = new XSSFWorkbook();

            // If "GeneralPlanDetails" exists, create a sheet
            if (cvs.has("GeneralPlanDetails")) {
                JSONObject generalPlanDetails = cvs.getJSONObject("GeneralPlanDetails");
                createSheetFromKeyValueObject(workbook, "GeneralPlanDetails", generalPlanDetails);
            }

            // If "Copay" (array) exists, create a sheet
            if (cvs.has("Copay")) {
                JSONArray copayArray = cvs.getJSONArray("Copay");
                createSheetFromKeyValueArray(workbook, "Copay", copayArray);
            }

            // Write to Excel
            fos = new FileOutputStream(excelFile);
            workbook.write(fos);
            System.out.println("Created Excel: " + excelFile.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Close resources
            if (workbook != null) {
                try { workbook.close(); } catch (Exception e) { /* ignore */ }
            }
            if (fos != null) {
                try { fos.close(); } catch (Exception e) { /* ignore */ }
            }
        }
    }

    /**
     * Creates a sheet from an object that has a "keyValue" array,
     * e.g. "GeneralPlanDetails": { "keyValue": [ { "attribute":"MobInd", "value":"N" } ] }
     */
    private static void createSheetFromKeyValueObject(Workbook workbook, String sheetName, JSONObject obj) {
        if (!obj.has("keyValue")) return;

        JSONArray keyValueArray = obj.getJSONArray("keyValue");

        // Gather attributes (column headers)
        Set<String> allAttributes = new LinkedHashSet<>();
        for (int i = 0; i < keyValueArray.length(); i++) {
            JSONObject kv = keyValueArray.getJSONObject(i);
            allAttributes.add(kv.getString("attribute"));
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // Header row
        Row header = sheet.createRow(0);
        List<String> attributes = new ArrayList<>(allAttributes);
        for (int i = 0; i < attributes.size(); i++) {
            header.createCell(i).setCellValue(attributes.get(i));
        }

        // Data row(s) - typically there's only one set of keyValue pairs,
        // but if you had more complex data, you'd adjust accordingly.
        Row row = sheet.createRow(1);
        for (int i = 0; i < keyValueArray.length(); i++) {
            JSONObject kv = keyValueArray.getJSONObject(i);
            String attr = kv.getString("attribute");
            String val = kv.getString("value");

            int colIndex = attributes.indexOf(attr);
            if (colIndex >= 0) {
                row.createCell(colIndex).setCellValue(val);
            }
        }

        // Auto-size columns
        for (int c = 0; c < attributes.size(); c++) {
            sheet.autoSizeColumn(c);
        }
    }

    /**
     * Creates a sheet from an array of objects, each having "keyValue" array.
     * e.g. "Copay": [
     *   { "keyValue": [ { "attribute":"CopayChannel","value":"RTL"} ] },
     *   { "keyValue": [ { "attribute":"CopayChannel","value":"MAIL"} ] }
     * ]
     */
    private static void createSheetFromKeyValueArray(Workbook workbook, String sheetName, JSONArray array) {
        // 1) Collect all possible attributes across all elements
        Set<String> allAttributes = new LinkedHashSet<>();
        List<JSONArray> keyValueArrays = new ArrayList<>();

        for (int i = 0; i < array.length(); i++) {
            JSONObject element = array.getJSONObject(i);
            JSONArray kvArr = element.optJSONArray("keyValue");
            if (kvArr != null) {
                keyValueArrays.add(kvArr);
                for (int j = 0; j < kvArr.length(); j++) {
                    JSONObject kv = kvArr.getJSONObject(j);
                    allAttributes.add(kv.getString("attribute"));
                }
            }
        }

        if (allAttributes.isEmpty()) {
            // No data to write
            return;
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // 2) Header row
        Row header = sheet.createRow(0);
        List<String> attributes = new ArrayList<>(allAttributes);
        for (int i = 0; i < attributes.size(); i++) {
            header.createCell(i).setCellValue(attributes.get(i));
        }

        // 3) Each element => one row
        int rowIndex = 1;
        for (JSONArray kvArr : keyValueArrays) {
            Row row = sheet.createRow(rowIndex++);
            for (int j = 0; j < kvArr.length(); j++) {
                JSONObject kv = kvArr.getJSONObject(j);
                String attr = kv.getString("attribute");
                String val = kv.getString("value");

                int colIndex = attributes.indexOf(attr);
                if (colIndex >= 0) {
                    row.createCell(colIndex).setCellValue(val);
                }
            }
        }

        // 4) Auto-size columns
        for (int c = 0; c < attributes.size(); c++) {
            sheet.autoSizeColumn(c);
        }
    }
}
