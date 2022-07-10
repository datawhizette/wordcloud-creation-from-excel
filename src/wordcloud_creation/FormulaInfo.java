package olytico;

/**
 *
 * @author Liliya
 */
public class FormulaInfo {

    private String sheetName;
    private Integer rowIndex;
    private Integer cellIndex;
    private String formula;

    FormulaInfo(String sheetName, Integer rowIndex, Integer cellIndex, String formula) {
        this.sheetName = sheetName;
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
        this.formula = formula;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(Integer rowIndex) {
        this.rowIndex = rowIndex;
    }

    public Integer getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(Integer cellIndex) {
        this.cellIndex = cellIndex;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }
}
