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

public class ExcelToJson {

    public static void main(String[] args) {

        String excelInputPath = "C:/Development/myCode/fromRequest/output.xlsx";
        String jsonOutputPath = "C:/Development/myCode/fromRequest/reconstructed.json";

        try (FileInputStream fis = new FileInputStream(excelInputPath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // We'll reconstruct the final JSON structure:
            // {
            //   "benefitRequest": {
            //     "transactionID": "...",
            //     "clientCode": "...",
            //     "data": ...,
            //     "dataSet": {
            //       "CVS": {
            //         "GeneralPlanDetails": { "keyValue": [ ... ] },
            //         "Copay": [ { "keyValue": [...] }, ... ]
            //       }
            //     }
            //   }
            // }
            JSONObject root = new JSONObject();
            JSONObject benefitRequest = new JSONObject();
            root.put("benefitRequest", benefitRequest);

            // Re-inject the fields we stored in memory during JsonToExcel
            // (If you are running this as a separate program, you'd need to pass them in.)
            benefitRequest.put("transactionID", JsonToExcel.transactionID);
            benefitRequest.put("clientCode", JsonToExcel.clientCode);

            // 'data' might be null or something else
            if (JsonToExcel.data == null) {
                benefitRequest.put("data", JSONObject.NULL);
            } else if (JsonToExcel.data instanceof JSONObject) {
                benefitRequest.put("data", (JSONObject)JsonToExcel.data);
            } else if (JsonToExcel.data instanceof JSONArray) {
                benefitRequest.put("data", (JSONArray)JsonToExcel.data);
            } else {
                // Just convert to string, or handle it as needed
                benefitRequest.put("data", JsonToExcel.data.toString());
            }

            // Now reconstruct "dataSet" → "CVS"
            JSONObject dataSet = new JSONObject();
            JSONObject cvs = new JSONObject();
            dataSet.put("CVS", cvs);
            benefitRequest.put("dataSet", dataSet);

            // Loop sheets to parse them back into JSON
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                if ("GeneralPlanDetails".equals(sheetName)) {
                    // Single object with "keyValue" array
                    JSONObject generalPlanDetails = new JSONObject();
                    parseKeyValueObjectSheet(sheet, generalPlanDetails);
                    cvs.put("GeneralPlanDetails", generalPlanDetails);
                }
                else if ("Copay".equals(sheetName)) {
                    // An array, each row becomes { "keyValue": [...] }
                    JSONArray copayArray = new JSONArray();
                    parseKeyValueArraySheet(sheet, copayArray);
                    cvs.put("Copay", copayArray);
                }
            }

            // Finally, write reconstructed JSON to file
            try (FileOutputStream fos = new FileOutputStream(new File(jsonOutputPath))) {
                fos.write(root.toString(4).getBytes());
            }
            System.out.println("Reconstructed JSON written to: " + jsonOutputPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * A sheet with 1 header row + 1 data row -> "keyValue": [...]
     * (Used for something like "GeneralPlanDetails".)
     */
    private static void parseKeyValueObjectSheet(Sheet sheet, JSONObject obj) {
        JSONArray keyValueArray = new JSONArray();

        Iterator<Row> rowIter = sheet.iterator();
        if (!rowIter.hasNext()) return; // no rows
        Row headerRow = rowIter.next();

        if (!rowIter.hasNext()) return; // no data row
        Row dataRow = rowIter.next();

        // For each cell in header row, read the data row
        int lastCellNum = headerRow.getPhysicalNumberOfCells();
        for (int c = 0; c < lastCellNum; c++) {
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
        obj.put("keyValue", keyValueArray);
    }

    /**
     * A sheet with 1 header row + multiple data rows -> each row = one object in the array
     * (Used for something like "Copay".)
     */
    private static void parseKeyValueArraySheet(Sheet sheet, JSONArray array) {
        Iterator<Row> rowIter = sheet.iterator();
        if (!rowIter.hasNext()) return;
        // Header row
        Row headerRow = rowIter.next();

        int headerCells = headerRow.getPhysicalNumberOfCells();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < headerCells; c++) {
            Cell cell = headerRow.getCell(c);
            String attribute = (cell == null) ? "" : cell.getStringCellValue();
            headers.add(attribute);
        }

        // Data rows
        while (rowIter.hasNext()) {
            Row dataRow = rowIter.next();
            JSONArray keyValueArray = new JSONArray();

            for (int c = 0; c < headers.size(); c++) {
                String attribute = headers.get(c);
                Cell dataCell = dataRow.getCell(c);
                String value = (dataCell == null) ? "" : dataCell.toString();

                JSONObject kv = new JSONObject();
                kv.put("attribute", attribute);
                kv.put("value", value);
                keyValueArray.put(kv);
            }

            // Each row => { "keyValue": [ ... ] }
            JSONObject copayObj = new JSONObject();
            copayObj.put("keyValue", keyValueArray);
            array.put(copayObj);
        }
    }
}
