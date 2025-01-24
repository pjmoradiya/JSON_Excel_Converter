package convertJSONtoJAVA;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileReader;
import java.io.FileWriter;
import java.nio.file.Files;
import java.nio.file.Paths;

public class requestCSVtoJSON {
	
	public static void main(String[] args) {
        String csvInputPath = "C:/Development/myCode/fromRequest/TestCopay.csv";
        String jsonOutputPath = "C:/Development/myCode/fromRequest/reversedOutput.json";

        try {
            FileReader fileReader = new FileReader(csvInputPath);
            CSVParser parser = CSVFormat.DEFAULT.withFirstRecordAsHeader().parse(fileReader);
            JSONArray keyValueArray = new JSONArray();

            for (CSVRecord record : parser) {
                JSONArray innerArray = new JSONArray();
                for (String header : parser.getHeaderMap().keySet()) {
                    JSONObject attributeValue = new JSONObject();
                    attributeValue.put("attribute", header);
                    attributeValue.put("value", record.get(header));
                    innerArray.put(attributeValue);
                }
                JSONObject obj = new JSONObject();
                obj.put("keyValue", innerArray);
                keyValueArray.put(obj);
            }

            JSONObject outputJson = new JSONObject().put("keyValue", keyValueArray);
            Files.write(Paths.get(jsonOutputPath), outputJson.toString(4).getBytes());
            System.out.println("JSON file has been created successfully");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}


