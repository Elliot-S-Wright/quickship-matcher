import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.Callable;

/**
 * Scans every .xlsx file in a BOMs folder and reports rows whose Part Number
 * (first column of the first sheet) matches a quick ship model number from the
 * SK Quickship report (second column). The 9th and 10th characters of the model
 * numbers are ignored for the comparison (they are wildcards like "**" in the
 * quick ship list, or door-position/hinge codes like "WL" in the BOMs).
 */
@Command(name = "quickship-matcher",
        description = "Finds quick ship model numbers in BOM Excel files.",
        mixinStandardHelpOptions = true,
        version = "quickship-matcher 1.0")
public class QuickShipMatcher implements Callable<Integer> {

    private static final DataFormatter FORMATTER = new DataFormatter();

    @Option(names = {"-q", "--quickship-file"}, required = true,
            description = "The SK Quickship report (.xlsx) with model numbers in column B")
    Path quickShipFile;

    @Option(names = {"-b", "--boms-dir"}, required = true,
            description = "Folder containing BOM .xlsx files (part numbers in column A of first sheet)")
    Path bomsDir;

    @Option(names = {"-o", "--out"}, defaultValue = "matches.csv",
            description = "CSV file to write matches to (default: ${DEFAULT-VALUE})")
    Path outFile;

    /** A matched row from a BOM file. */
    record Match(String partNumber, String description) {
    }

    public static void main(String[] args) {
        System.exit(new CommandLine(new QuickShipMatcher()).execute(args));
    }

    @Override
    public Integer call() throws IOException {
        if (!Files.isRegularFile(quickShipFile)) {
            throw new IllegalArgumentException("Quickship file not found: " + quickShipFile.toAbsolutePath());
        }
        if (!Files.isDirectory(bomsDir)) {
            throw new IllegalArgumentException("BOMs folder not found: " + bomsDir.toAbsolutePath());
        }

        Set<String> quickShipKeys = loadQuickShipModels(quickShipFile);
        if (quickShipKeys.isEmpty()) {
            throw new IllegalStateException("No model numbers found in column B of "
                    + quickShipFile.getFileName() + " - has the report layout changed?");
        }
        System.out.println("Loaded " + quickShipKeys.size() + " quick ship model numbers.");
        System.out.println();

        List<Path> bomFiles = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(bomsDir, "*.xlsx")) {
            for (Path p : stream) {
                // Skip "~$..." lock files Excel creates while a workbook is open.
                if (!p.getFileName().toString().startsWith("~$")) {
                    bomFiles.add(p);
                }
            }
        }
        if (bomFiles.isEmpty()) {
            throw new IllegalStateException("No .xlsx files found in " + bomsDir.toAbsolutePath());
        }
        bomFiles.sort(null);

        int matchedFiles = 0;
        int failedFiles = 0;
        List<String> csvLines = new ArrayList<>();
        csvLines.add("file,part_number,description");
        for (Path bomFile : bomFiles) {
            try {
                List<Match> matches = searchBomFile(bomFile, quickShipKeys);
                if (!matches.isEmpty()) {
                    matchedFiles++;
                    for (Match match : matches) {
                        csvLines.add(toCsvRow(bomFile.getFileName().toString(),
                                match.partNumber(), match.description()));
                    }
                }
            } catch (Exception e) {
                failedFiles++;
                System.err.println("Failed to read " + bomFile.getFileName() + ": " + e.getMessage());
            }
        }

        Files.write(outFile, csvLines);
        System.out.println("Wrote " + (csvLines.size() - 1) + " matches to " + outFile.toAbsolutePath());
        System.out.println("Done. Searched " + bomFiles.size() + " files, found matches in "
                + matchedFiles + ", failed to read " + failedFiles + ".");
        return failedFiles == 0 ? 0 : 2;
    }

    /** Reads the second column of the quick ship report and returns the normalized model numbers. */
    static Set<String> loadQuickShipModels(Path file) throws IOException {
        Set<String> keys = new HashSet<>();
        try (InputStream in = Files.newInputStream(file);
             Workbook workbook = new XSSFWorkbook(in)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String value = cellText(row, 1);
                // Skip blanks and header rows; real model numbers start with "QC"/"QF".
                if (value.isEmpty() || value.equalsIgnoreCase("QuickShip")) {
                    continue;
                }
                keys.add(normalize(value));
            }
        }
        return keys;
    }

    /**
     * Searches the first sheet of a BOM file and returns the rows whose part number
     * (column A) matches a quick ship model, with their descriptions (column B).
     */
    static List<Match> searchBomFile(Path file, Set<String> quickShipKeys) throws IOException {
        List<Match> matches = new ArrayList<>();
        try (InputStream in = Files.newInputStream(file);
             Workbook workbook = new XSSFWorkbook(in)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String partNumber = cellText(row, 0);
                if (partNumber.isEmpty()) {
                    continue;
                }
                if (quickShipKeys.contains(normalize(partNumber))) {
                    matches.add(new Match(partNumber, cellText(row, 1)));
                }
            }
        }
        return matches;
    }

    /** Builds one CSV row, escaping fields per RFC 4180. */
    static String toCsvRow(String... fields) {
        StringBuilder row = new StringBuilder();
        for (int i = 0; i < fields.length; i++) {
            if (i > 0) {
                row.append(',');
            }
            row.append(csvEscape(fields[i]));
        }
        return row.toString();
    }

    /** Quotes a field if it contains commas, quotes, or newlines; doubles embedded quotes. */
    static String csvEscape(String field) {
        if (field.contains(",") || field.contains("\"") || field.contains("\n") || field.contains("\r")) {
            return '"' + field.replace("\"", "\"\"") + '"';
        }
        return field;
    }

    /** Removes the 9th and 10th characters so e.g. QC081277WLN matches QC081277**N. */
    static String normalize(String value) {
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
