import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.Ref3DPtg;
import org.apache.poi.ss.formula.ptg.Area3DPtg;
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
        Map<String, List<String>> map = new Hashtable<>();
        StringBuilder s = new StringBuilder("digraph G {\n");
        for (int i = 0; i < wb.getNumberOfSheets(); i++)
        {
            Sheet sheet = wb.getSheetAt(i);
            sheet.forEach(r -> r.forEach(c -> {
                List<String> children = getCellReferences(c);
                if (children != null) map.put(getSheetPrefix(c) + new CellReference(c).formatAsString(), children);
            }));
        }
        map.forEach((key, value) -> value.forEach(cell -> s.append("\"" + key + "\" -> \"" + cell + "\";\n")));
        s.append("}");
        return s.toString();
    }

    /**
     * @param c A cell.
     * @return List of cells referenced by this location/range, like [A4], [J7], or [A4, A5, A6].
     */
    private List<String> getCellReferences(Cell c) {
        Ptg[] ptgs;
        try {
            ptgs = FormulaParser.parse(c.getCellFormula(), xssfew, c.getCellType(), 0);
        } catch (java.lang.IllegalStateException e) {
            if (e.toString().contains("text cell") || e.toString().contains("numeric cell") ||
                    e.toString().contains("blank cell")) return null;
            else throw e;
        }

        List<String> returning = new ArrayList<>();
        for (Ptg p : ptgs) {
            String[] split = p.getClass().toString().split("\\.");
            switch (split[split.length - 1]) {
                case "RefPtg":
                    returning.add(getSheetPrefix(c) + p.toFormulaString());
                    break;
                case "AreaPtg":
                    getRange(p.toFormulaString()).forEach(str -> returning.add(getSheetPrefix(c) + str));
                    break;
                case "Ref3DPtg":
                    Ref3DPtg p3d = (Ref3DPtg) p;
                    String s = p3d.toFormulaString(xssfew);
                    returning.add(s);
                    break;
                case "Area3DPtg":
                    Area3DPtg a3d = (Area3DPtg) p;
                    getRange(a3d.format2DRefAsString()).forEach(str -> returning.add(getSheetPrefix(c) + str));
                    break;
                default:
                    break;
            }
        }
        return returning;
    }

    /**
     * @param c a cell.
     * @return The cell's prefix, probably something like "'Sheet #1'!"
     */
    private static String getSheetPrefix(Cell c) {
        return "'" + c.getSheet().getSheetName() + "'!";
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
}
