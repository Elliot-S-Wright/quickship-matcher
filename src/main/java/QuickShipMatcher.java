import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Scans every .xlsx file in the BOMs folder and reports rows whose Part Number
 * (first column of the first sheet) matches a quick ship model number from the
 * SK Quickship report (second column). The 9th and 10th characters of the model
 * numbers are ignored for the comparison (they are wildcards like "**" in the
 * quick ship list, or door-option codes like "WL" in the BOMs).
 */
public class QuickShipMatcher {

    private static final DataFormatter FORMATTER = new DataFormatter();

    public static void main(String[] args) throws IOException {
        Path baseDir = Path.of(args.length > 0 ? args[0] : ".");
        Path quickShipFile = baseDir.resolve("SK Quickship Missing Cost Report.xlsx");
        Path bomsDir = baseDir.resolve("BOMs");

        Set<String> quickShipKeys = loadQuickShipModels(quickShipFile);
        System.out.println("Loaded " + quickShipKeys.size() + " quick ship model numbers.");
        System.out.println();

        List<Path> bomFiles = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(bomsDir, "*.xlsx")) {
            stream.forEach(bomFiles::add);
        }
        bomFiles.sort(null);

        int matchedFiles = 0;
        for (Path bomFile : bomFiles) {
            if (searchBomFile(bomFile, quickShipKeys)) {
                matchedFiles++;
            }
        }

        System.out.println();
        System.out.println("Done. Searched " + bomFiles.size() + " files, found matches in " + matchedFiles + ".");
    }

    /** Reads the second column of the quick ship report and returns the normalized model numbers. */
    private static Set<String> loadQuickShipModels(Path file) throws IOException {
        Set<String> keys = new HashSet<>();
        try (InputStream in = Files.newInputStream(file);
             Workbook workbook = new XSSFWorkbook(in)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String value = cellText(row, 1);
                // Skip blanks and header rows; real model numbers start with "QC".
                if (value.isEmpty() || value.equalsIgnoreCase("QuickShip")) {
                    continue;
                }
                keys.add(normalize(value));
            }
        }
        return keys;
    }

    /**
     * Searches the first sheet of a BOM file for part numbers matching a quick ship
     * model. Prints any matches and returns whether at least one was found.
     */
    private static boolean searchBomFile(Path file, Set<String> quickShipKeys) {
        boolean found = false;
        try (InputStream in = Files.newInputStream(file);
             Workbook workbook = new XSSFWorkbook(in)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String partNumber = cellText(row, 0);
                if (partNumber.isEmpty()) {
                    continue;
                }
                if (quickShipKeys.contains(normalize(partNumber))) {
                    String description = cellText(row, 1).trim();
                    System.out.println(file.getFileName() + " | " + partNumber + " | " + description);
                    found = true;
                }
            }
        } catch (Exception e) {
            System.err.println("Failed to read " + file.getFileName() + ": " + e.getMessage());
        }
        return found;
    }

    /** Removes the 9th and 10th characters so e.g. QC081277WLN matches QC060682**N. */
    private static String normalize(String value) {
        String v = value.trim().toUpperCase();
        if (v.length() >= 10) {
            v = v.substring(0, 8) + v.substring(10);
        }
        return v;
    }

    private static String cellText(Row row, int columnIndex) {
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(columnIndex);
        return cell == null ? "" : FORMATTER.formatCellValue(cell).trim();
    }
}
