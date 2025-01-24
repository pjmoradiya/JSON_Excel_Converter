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

public class JsonToExcel {

    // We'll store these benefitRequest fields in memory so we can restore them later.
    // In a real application, you might store them in a DB, pass them to the next step, or
    // put them in a hidden sheet if you want them inside the workbook.
    public static String transactionID;
    public static String clientCode;
    public static Object data; // could be null or some object

    public static void main(String[] args) {

        String jsonInputPath = "C:/Development/myCode/fromRequest/sample.json";
        String excelOutputPath = "C:/Development/myCode/fromRequest/output.xlsx";

        try {
            // 1) Read the JSON text from file
            String jsonString = new String(Files.readAllBytes(Paths.get(jsonInputPath)));
            JSONObject root = new JSONObject(jsonString);

            // 2) Grab "benefitRequest" (one object)
            JSONObject benefitRequest = root.optJSONObject("benefitRequest");
            if (benefitRequest == null) {
                System.out.println("No benefitRequest object found!");
                return;
            }

            // Extract top-level fields from benefitRequest that we do NOT want in Excel,
            // but do want to preserve for later.
            transactionID = benefitRequest.optString("transactionID");
            clientCode = benefitRequest.optString("clientCode");
            // 'data' could be null or any value (in your example, it's null).
            // We'll just store whatever it is.
            data = benefitRequest.opt("data");

            // Now let's create the workbook for only the details we want in Excel.
            Workbook workbook = new XSSFWorkbook();

            // 3) dataSet → CVS → GeneralPlanDetails
            JSONObject dataSet = benefitRequest.optJSONObject("dataSet");
            if (dataSet != null) {
                JSONObject cvs = dataSet.optJSONObject("CVS");
                if (cvs != null) {
                    // a) GeneralPlanDetails
                    if (cvs.has("GeneralPlanDetails")) {
                        JSONObject generalPlanDetails = cvs.getJSONObject("GeneralPlanDetails");
                        createSheetFromKeyValueObject(workbook, "GeneralPlanDetails", generalPlanDetails);
                    }

                    // b) Copay (array)
                    if (cvs.has("Copay")) {
                        JSONArray copayArray = cvs.getJSONArray("Copay");
                        createSheetFromKeyValueArray(workbook, "Copay", copayArray);
                    }
                }
            }

            // 4) Write the workbook to an Excel file
            try (FileOutputStream fos = new FileOutputStream(new File(excelOutputPath))) {
                workbook.write(fos);
            }
            workbook.close();

            System.out.println("Excel file created: " + excelOutputPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Creates a sheet from an object that has a "keyValue" array
     * (e.g. "GeneralPlanDetails": {"keyValue": [ ... ]}).
     */
    private static void createSheetFromKeyValueObject(Workbook workbook, String sheetName, JSONObject obj) {
        if (!obj.has("keyValue")) return;

        // "keyValue" is an array of { "attribute": ..., "value": ... }
        JSONArray keyValueArray = obj.getJSONArray("keyValue");

        // Gather attributes
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

        // We'll assume there's usually only one row for "GeneralPlanDetails".
        // But let's handle the possibility of multiple anyway.
        for (int r = 0; r < 1 /* or keyValueArray.length() for multi-row? */; r++) {
            Row row = sheet.createRow(r + 1); // data row
            for (int i = 0; i < keyValueArray.length(); i++) {
                JSONObject kv = keyValueArray.getJSONObject(i);
                String attr = kv.getString("attribute");
                String val = kv.getString("value");

                int colIndex = attributes.indexOf(attr);
                if (colIndex >= 0) {
                    row.createCell(colIndex).setCellValue(val);
                }
            }
        }

        // Auto-size columns
        for (int c = 0; c < attributes.size(); c++) {
            sheet.autoSizeColumn(c);
        }
    }

    /**
     * Creates a sheet from an array, each element having {"keyValue": [...]}.
     * (e.g. "Copay": [ { "keyValue": [...] }, { "keyValue": [...] } ])
     */
    private static void createSheetFromKeyValueArray(Workbook workbook, String sheetName, JSONArray array) {
        // Collect all possible attributes across all objects
        Set<String> allAttributes = new LinkedHashSet<>();
        List<JSONArray> keyValueArrays = new ArrayList<>();

        for (int i = 0; i < array.length(); i++) {
            JSONObject element = array.getJSONObject(i);
            JSONArray kvArray = element.optJSONArray("keyValue");
            if (kvArray != null) {
                keyValueArrays.add(kvArray);
                for (int j = 0; j < kvArray.length(); j++) {
                    JSONObject kv = kvArray.getJSONObject(j);
                    allAttributes.add(kv.getString("attribute"));
                }
            }
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // Header row
        Row header = sheet.createRow(0);
        List<String> attributes = new ArrayList<>(allAttributes);
        for (int i = 0; i < attributes.size(); i++) {
            header.createCell(i).setCellValue(attributes.get(i));
        }

        // Each element in the array => one row
        int rowIndex = 1;
        for (JSONArray kvArray : keyValueArrays) {
            Row row = sheet.createRow(rowIndex++);
            for (int j = 0; j < kvArray.length(); j++) {
                JSONObject kv = kvArray.getJSONObject(j);
                String attr = kv.getString("attribute");
                String val = kv.getString("value");

                int colIndex = attributes.indexOf(attr);
                if (colIndex >= 0) {
                    row.createCell(colIndex).setCellValue(val);
                }
            }
        }

        // Auto-size columns
        for (int c = 0; c < attributes.size(); c++) {
            sheet.autoSizeColumn(c);
        }
    }
}
