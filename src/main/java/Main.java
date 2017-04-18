import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;

public class Main {
    private Workbook wb;
    private XSSFEvaluationWorkbook xssfew;

    /** Instantiates a Main class creating an XSSF workbook from the given filename.
     * @param fileName should end in .xlsx or you're gonna need to modify the code to use HSSF Workbooks and Evaluators.
     * @throws Exception if spreadsheet was invalid format or not found.
     */
    private Main(String fileName) throws Exception {
        wb = WorkbookFactory.create(new FileInputStream(fileName));
        xssfew = XSSFEvaluationWorkbook.create((XSSFWorkbook) wb);
    }

    public static void main(String... args) throws Exception {
        System.out.println(new Main(args[0]).translateSpreadsheetToDot());
    }

    /** @return Cell dependencies for the active sheet translated to graph-viz dot language */
    private String translateSpreadsheetToDot() {
        Sheet sheet = wb.getSheetAt(wb.getActiveSheetIndex());
        Map<String, List<String>> map = new Hashtable<>();

        sheet.forEach(x -> x.forEach(c -> {
            List<String> children = getCellReferences(c);
            if (children != null) map.put(new CellReference(c).formatAsString(), children);
        }));

        StringBuilder s = new StringBuilder("digraph G {\n");
        map.forEach((key, value) -> value.forEach(cell -> s.append(key + " -> " + cell + ";\n")));
        s.append("}");

        return s.toString();
    }

    /**
     * @param c A cell.
     * @return List of cells represented by this location/range, like [A4], [J7], or [A4, A5, A6].
     */
    private List<String> getCellReferences(Cell c) {
        Ptg[] ptgs;
        try {
            ptgs = FormulaParser.parse(c.getCellFormula(), xssfew, c.getCellType(), 0);
        } catch (Exception e) {
            return null;
        }
        List<String> returning = new ArrayList<>();
        for (Ptg p : ptgs) {
            if (isRef(p)) {
                String s = p.toFormulaString();
                if (s.contains(":")) returning.addAll(getRange(s));
                else returning.add(s);
            }
        }
        return returning;
    }


    /**
     * @param s String encoding of a cell range like "A6:A8"
     * @return List of String encodings of cells pointed to, like [A6, A7, A8]
     */
    private static List<String> getRange(String s) {
        List<String> returning = new ArrayList<>();
        String cells[] = s.split(":");
        CellReference first = new CellReference(cells[0]);
        CellReference last = new CellReference(cells[1]);

        for (int i = first.getRow(); i <= last.getRow(); i++) {
            for (int j = first.getCol(); j <= last.getCol(); j++) {
                returning.add(new CellReference(i, j).formatAsString());
            }
        }
        return returning;
    }

    /** @return Whether or not p was a cell reference. */
    private static boolean isRef(Ptg p) {
        String[] references = {"org.apache.poi.ss.formula.ptg.AreaPtg", "org.apache.poi.ss.formula.ptg.RefPtg"};
        for (String ref : references) {
            try {
                if (p.getClass().equals(Class.forName(ref))) {
                    return true;
                }
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
                System.out.println(p.getClass());
            }
        }
        return false;
    }
}
