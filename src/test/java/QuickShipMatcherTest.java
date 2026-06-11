import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class QuickShipMatcherTest {

    @Nested
    class Normalize {

        @Test
        void stripsNinthAndTenthCharacters() {
            assertEquals("QC060877F", QuickShipMatcher.normalize("QC060877WLF"));
            assertEquals("QC060877F", QuickShipMatcher.normalize("QC060877**F"));
        }

        @Test
        void wildcardAndConcreteModelNumbersMatch() {
            assertEquals(QuickShipMatcher.normalize("QC081277**N"),
                    QuickShipMatcher.normalize("QC081277WLN"));
        }

        @Test
        void isCaseAndWhitespaceInsensitive() {
            assertEquals("QC060877F", QuickShipMatcher.normalize(" qc060877wlf "));
        }

        @Test
        void shortPartNumbersPassThroughUnchanged() {
            assertEquals("3587", QuickShipMatcher.normalize("3587"));
        }

        @Test
        void differentSizeModelsDoNotCollide() {
            // Only the door codes (chars 9-10) are wildcards; dimensions must still differ.
            assertEquals("QC060877F", QuickShipMatcher.normalize("QC060877WLF"));
            assertEquals("QC081277F", QuickShipMatcher.normalize("QC081277WLF"));
        }
    }

    @Nested
    class CsvOutput {

        @Test
        void plainFieldsArePassedThrough() {
            assertEquals("bom.xlsx,QC060682WLN,Simple description",
                    QuickShipMatcher.toCsvRow("bom.xlsx", "QC060682WLN", "Simple description"));
        }

        @Test
        void fieldsWithCommasAreQuoted() {
            // Real descriptions contain commas, e.g. "Quick Ship Indoor Walk-In, 6' W x 6' L..."
            assertEquals("bom.xlsx,QC060682WLN,\"Quick Ship Indoor Walk-In, 6' W x 6' L\"",
                    QuickShipMatcher.toCsvRow("bom.xlsx", "QC060682WLN", "Quick Ship Indoor Walk-In, 6' W x 6' L"));
        }

        @Test
        void embeddedQuotesAreDoubled() {
            assertEquals("\"NSF Vinyl Floor Screed 72\"\"\"",
                    QuickShipMatcher.csvEscape("NSF Vinyl Floor Screed 72\""));
        }

        @Test
        void fieldsWithNewlinesAreQuoted() {
            assertEquals("\"Part\nNumber\"", QuickShipMatcher.csvEscape("Part\nNumber"));
        }
    }

    @Nested
    class ExcelReading {

        @TempDir
        Path tempDir;

        @Test
        void loadsModelNumbersFromSecondColumnSkippingHeaders() throws IOException {
            Path file = tempDir.resolve("quickship.xlsx");
            writeWorkbook(file,
                    row(null, null),
                    row(null, "QuickShip"),          // header row, must be skipped
                    row(null, "QC060682**N"),
                    row(null, "QF081277**F"));

            Set<String> keys = QuickShipMatcher.loadQuickShipModels(file);

            assertEquals(Set.of("QC060682N", "QF081277F"), keys);
        }

        @Test
        void findsMatchingPartNumbersWithDescriptions() throws IOException {
            Path bom = tempDir.resolve("bom.xlsx");
            writeWorkbook(bom,
                    row("Part\nNumber", "Description"),
                    row("QC060682WLN", "Quick Ship Walk-In Cooler"),
                    row("3587", "ALLEN WRENCH"),
                    row("QF999999XXX", "Not a quick ship model"));

            List<QuickShipMatcher.Match> matches =
                    QuickShipMatcher.searchBomFile(bom, Set.of("QC060682N"));

            assertEquals(List.of(new QuickShipMatcher.Match("QC060682WLN", "Quick Ship Walk-In Cooler")),
                    matches);
        }

        @Test
        void returnsNoMatchesWhenNothingMatches() throws IOException {
            Path bom = tempDir.resolve("bom.xlsx");
            writeWorkbook(bom,
                    row("Part\nNumber", "Description"),
                    row("3587", "ALLEN WRENCH"));

            assertTrue(QuickShipMatcher.searchBomFile(bom, Set.of("QC060682N")).isEmpty());
        }

        private static String[] row(String colA, String colB) {
            return new String[]{colA, colB};
        }

        private static void writeWorkbook(Path file, String[]... rows) throws IOException {
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Sheet1");
                for (int r = 0; r < rows.length; r++) {
                    Row row = sheet.createRow(r);
                    for (int c = 0; c < rows[r].length; c++) {
                        if (rows[r][c] != null) {
                            row.createCell(c).setCellValue(rows[r][c]);
                        }
                    }
                }
                try (OutputStream out = Files.newOutputStream(file)) {
                    workbook.write(out);
                }
            }
        }
    }
}
