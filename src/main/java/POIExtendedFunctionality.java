import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * This Class contains some Functionality,
 * that you might need while working with Excel Files using Apache POI
 */
public class POIExtendedFunctionality {

    /**
     * Insert Column at the specified index
     * NOTE: Formulas will not be updated or re-evaluated
     * @param workbook The workbook (Actually is the Excel file) that contains the Sheet you want to insert the new column into.
     * @param sheetIndex Sheet Index in that Workbook (Starts by 0)
     * @param newColumnIndex Index at which you want the new Column to be inserted.
     * @throws NullPointerException
     */
    public static void insertNewColumn(Workbook workbook, int sheetIndex, int newColumnIndex) throws NullPointerException {
        if (workbook == null)
            throw new NullPointerException("Error: Workbook can not be NULL!");

        Sheet sheet = workbook.getSheetAt(sheetIndex);

        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.clearAllCachedResultValues();

        int nrRows = sheet.getLastRowNum() + 1;
        int nrCols = sheet.getRow(2).getLastCellNum();

        // Iterate through all columns
        for (int i = 0; i < nrRows; i++) {
            Row row = sheet.getRow(i);

            if (row == null) {
                continue;
            }

            // move columns to right
            for (int col = nrCols; col > newColumnIndex; col--) {
                Cell rightCell = row.getCell(col);

                if (rightCell != null)
                    row.removeCell(rightCell);

                Cell leftCell = row.getCell(col - 1);
                if (leftCell != null) {
                    Cell newCell = row.createCell(col, leftCell.getCellType());
                    cloneCellType(newCell, leftCell);
                }
            }

            // Delete old column
            Cell currentEmptyWeekCell = row.getCell(newColumnIndex);
            if (currentEmptyWeekCell != null) {
                row.removeCell(currentEmptyWeekCell);
            }
            // Creating the new Cell at the new Column position
            row.createCell(newColumnIndex, CellType.BLANK);
        }

        // Adjust the column widths
        for (int col = nrCols; col > newColumnIndex; col--) {
            sheet.setColumnWidth(col, sheet.getColumnWidth(col - 1));
        }

        XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
    }

    /**
     * Clones Old-Cell type to the New-Cell
     * @param newCell cell without type
     * @param oldCell cell with type to be cloned
     */
    public static void cloneCellType(Cell newCell, Cell oldCell) {
        newCell.setCellComment(oldCell.getCellComment());
        newCell.setCellStyle(oldCell.getCellStyle());
        switch (oldCell.getCellType()) {
            case BOOLEAN: {
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            }
            case _NONE:
                break;
            case NUMERIC: {
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            }
            case STRING: {
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            }
            case ERROR: {
                newCell.setCellValue(oldCell.getErrorCellValue());
                break;
            }
            case FORMULA: {
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            }
            case BLANK:
                break;
        }
    }

    /**
     * Does a "Base 26" - Base26E transformation on the given index, to obtain the alphabet representation.
     * The transformation is not exactly Base26, since the factor for each degree of power (besides first)
     * is represented as "1 less".
     * <p>
     * Ex:
     * 25 -> Z
     * 26 -> BA (in Base26) -> AA (in Excel)
     * 27 -> BB (in Base26) -> AB (in Excel)
     * (we have B instead of A for degree of power 1)
     * So a normal 'AACAAA' in Base26 is 'BBDBBA' in Base26E.
     * <p>
     * This is how excel identifies columns in formulas.
     *
     * @param columnIndex Numeric index of column
     * @return Alphabetic index of column as it is represented in Excel
     */
    public static String getColumnAlphabeticIndex(int columnIndex) {
        StringBuilder alphabeticIndex = new StringBuilder();
        char LETTERS_IN_EN_ALPHABET = 26;
        char A_LETTER = 65;

        while (columnIndex >= 0) {
            if (columnIndex == 0) {
                alphabeticIndex.append(A_LETTER);
                break;
            }
            char code = (char) (columnIndex % LETTERS_IN_EN_ALPHABET);
            char letter = (char) (code + A_LETTER);
            alphabeticIndex.append(letter);

            columnIndex /= LETTERS_IN_EN_ALPHABET;
            columnIndex -= 1;
        }

        return alphabeticIndex.reverse().toString();
    }

    /**
     * Save changes on file
     */
    public static void saveChanges(Workbook workbook, String path) {
        // TODO: use one path
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(path);
            workbook.setForceFormulaRecalculation(true);
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
