package csv;

import java.io.*;
import java.util.*;

/**
 * CsvConverter is a class that processes a CSV file and converts it into a map of records.
 * Each record is represented as a list of values.
 *
 * @param <T> The type of values in the record list
 */
public class CsvReader<T> {

    private final Map<Integer, List<T>> records = new HashMap<>();
    private final int[] selectedColumns = {0, 1, 49, 50, 20, 18, 51, 12, 54, 52, 53};
    private final File file;

    public CsvReader(String filePath) {
        this.file = new File(filePath);
        processCsvFile();
    }

    public List<List<T>> getRecordsAsList() {
        return new ArrayList<>(records.values());
    }

    private void processCsvFile() {
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String csvRow;
            int rowIndex = 0;
            while ((csvRow = br.readLine()) != null) {
                List<String> rowValues = new ArrayList<>(Arrays.stream(csvRow.split(",")).toList());
                if (rowValues.size() < 54 | rowValues.contains("DUPE")) {
                    continue;
                }
                List<String> sortedRowValues = new ArrayList<>();
                for (int columnIndex : selectedColumns) {
                    sortedRowValues.add(rowValues.get(columnIndex));
                }
                sortedRowValues.add(7, "(Not Given)");
                sortedRowValues.add(12, "(Not Given)");
                records.put(rowIndex, (List<T>) sortedRowValues);
                rowIndex++;
            }
        } catch (IOException e) {
            throw new RuntimeException("Error while reading file", e);
        }
    }

}
