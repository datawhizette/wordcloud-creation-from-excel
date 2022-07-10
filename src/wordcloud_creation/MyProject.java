/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package olytico;

import com.sun.javafx.css.StyleCache;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import org.apache.poi.hssf.record.NameRecord;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataBarFormatting;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting.IconSet;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Liliya A
 */
public class MyProject {
    private static final String[] titles = {
            "Sport", "Tech", "Business", "Media", "Youth", "CEO/Founders", "Music", "Blogger" };
    private static final String[] colors = {
            "sport", "tech", "business", "media", "youth", "ceo", "music", "blogger" };

    private static final String[] letters = {
            "A", "B", "C", "D", "E", "F", "G", "H" };

    public static void main(String[] args) throws IOException {
        String file = "Vodafone Bio Segments November 2016.xlsx";
        // String file = "bios.xlsx";
        FileInputStream in = new FileInputStream(file);
        Workbook wb_main = new XSSFWorkbook(in);

        String file1 = "Vodafone Bio Keywords and Categories.xlsx";
        // String file1 = "keywords.xlsx";
        FileInputStream in1 = new FileInputStream(file1);
        Workbook wb_kw = new XSSFWorkbook(in1);

        Workbook wb_bios = new XSSFWorkbook();
        Workbook wb_sport = new XSSFWorkbook();
        Workbook wb_tech = new XSSFWorkbook();
        Workbook wb_business = new XSSFWorkbook();
        Workbook wb_media = new XSSFWorkbook();
        Workbook wb_youth = new XSSFWorkbook();
        Workbook wb_ceo = new XSSFWorkbook();
        Workbook wb_music = new XSSFWorkbook();
        Workbook wb_bloggers = new XSSFWorkbook();

        Sheet cbios = wb_bios.createSheet("All bios");
        Sheet csport = wb_sport.createSheet("Sport");
        Sheet ctech = wb_tech.createSheet("Tech");
        Sheet cbusiness = wb_business.createSheet("Business");
        Sheet cmedia = wb_media.createSheet("Media");
        Sheet cyouth = wb_youth.createSheet("Youth");
        Sheet cceo = wb_ceo.createSheet("CEO&Founders");
        Sheet cmusic = wb_music.createSheet("Music");
        Sheet cbloggers = wb_bloggers.createSheet("Blogger");

        boolean copyStyle = true;
        List<CellStyle> styleMap = (copyStyle) ? new ArrayList<>() : null;
        Map<String, CellStyle> styles = createStyles(wb_main);

        Sheet bios = wb_main.getSheet("All bios");
        Sheet sport = wb_main.createSheet("Sport");
        Sheet tech = wb_main.createSheet("Tech");
        Sheet business = wb_main.createSheet("Business");
        Sheet media = wb_main.createSheet("Media");
        Sheet youth = wb_main.createSheet("Youth");
        Sheet ceo = wb_main.createSheet("CEO&Founders");
        Sheet music = wb_main.createSheet("Music");
        Sheet bloggers = wb_main.createSheet("Blogger");

        Sheet keywords = wb_kw.getSheet("Keywords");

        Row headerRow = bios.getRow(0);

        for (int i = 0; i < titles.length; i++) {
            Cell cell = headerRow.createCell(i + 15);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(styles.get(colors[i]));
        }

        copyRow(bios, sport, headerRow, sport.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, tech.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, business.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, media.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, youth.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, ceo.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, music.createRow(0), styleMap);
        copyRow(bios, sport, headerRow, bloggers.createRow(0), styleMap);

        int rowEnd = bios.getLastRowNum();
        System.out.println("Number of bios " + rowEnd);

        Iterator rowIter = keywords.rowIterator(); // to get the number of keywords for each category
        Row r = (Row) rowIter.next();
        short lastCellNum = r.getLastCellNum();
        int[] dataCount = new int[lastCellNum];
        int col = 0;
        rowIter = keywords.rowIterator();
        while (rowIter.hasNext()) {
            Iterator cellIter = ((Row) rowIter.next()).cellIterator();
            while (cellIter.hasNext()) {
                Cell cell = (Cell) cellIter.next();
                col = cell.getColumnIndex();
                dataCount[col] += 1;

            }
        }

        for (int x = 0; x < dataCount.length; x++) { // iterate over columns
            for (int i = 0; i < dataCount[x] - 1; i++) // iterate over words in one category column
            {
                String keyword = letters[x] + (i + 2);
                CellReference cellReference = new CellReference(keyword);
                Row row = keywords.getRow(cellReference.getRow());
                Cell cell = row.getCell(cellReference.getCol());

                for (int j = 2; j < rowEnd + 2; j++) // iterate over posts
                {
                    String post = "B" + (j);
                    CellReference cellReference2 = new CellReference(post);
                    Row row2 = bios.getRow(cellReference2.getRow());
                    Cell cell2 = row2.getCell(cellReference2.getCol(), row2.CREATE_NULL_AS_BLANK);

                    if (cell.getStringCellValue().length() < 3) {
                        List<String> txt = Arrays.asList(cell2.toString().toLowerCase().split(" "));

                        if (txt.contains(cell.getStringCellValue().toLowerCase())) {
                            if (x == 0) // rule for each category
                            {
                                String ones = "Q" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                cell2.setCellStyle(styles.get("tech"));
                            } else if (x == 1) // rule for each category
                            {
                                String ones = "P" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(2, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("sport"));
                            } else if (x == 2) // rule for each category
                            {
                                String ones = "V" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(3, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("music"));
                            } else if (x == 3) // rule for each category
                            {
                                String ones = "R" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(4, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("business"));
                            } else if (x == 4) // rule for each category
                            {
                                String ones = "T" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(5, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("youth"));
                            } else if (x == 5) // rule for each category
                            {
                                String ones = "U" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(6, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("ceo"));
                            } else if (x == 6) // rule for each category
                            {
                                String ones = "S" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(7, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("media"));
                            } else if (x == 7) // rule for each category
                            {
                                String ones = "W" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(8, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("blogger"));
                            }
                        }
                    } else {
                        String txt;
                        txt = cell2.toString().toLowerCase();

                        if (txt.contains(cell.getStringCellValue().toLowerCase())) {
                            if (x == 0) // rule for each category
                            {
                                String ones = "Q" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                cell2.setCellStyle(styles.get("tech"));
                            } else if (x == 1) // rule for each category
                            {
                                String ones = "P" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(2, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("sport"));
                            } else if (x == 2) // rule for each category
                            {
                                String ones = "V" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(3, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("music"));
                            } else if (x == 3) // rule for each category
                            {
                                String ones = "R" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(4, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("business"));
                            } else if (x == 4) // rule for each category
                            {
                                String ones = "T" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(5, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("youth"));
                            } else if (x == 5) // rule for each category
                            {
                                String ones = "U" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(6, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("ceo"));
                            } else if (x == 6) // rule for each category
                            {
                                String ones = "S" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(7, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("media"));
                            } else if (x == 7) // rule for each category
                            {
                                String ones = "W" + j;
                                CellReference cellReference3 = new CellReference(ones);
                                Row row3 = bios.getRow(cellReference3.getRow());
                                Cell cell3 = row3.createCell(cellReference3.getCol());
                                cell3.setCellValue(1);
                                Cell cellAdd = row2.getCell(8, row2.CREATE_NULL_AS_BLANK);
                                cellAdd.setCellStyle(styles.get("blogger"));
                            }
                        }
                    }
                }
            }
        }
        DataFormatter formatter = new DataFormatter(); // transfer rows to other sheet
        int S_number = 1;
        int T_number = 1;
        int B_number = 1;
        int M_number = 1;
        int Y_number = 1;
        int C_number = 1;
        int Mu_number = 1;
        int Bl_number = 1;

        for (int rowNum = 1; rowNum < rowEnd + 2; rowNum++) // iterate over posts
        {
            Row row3 = bios.getRow(rowNum);

            if (row3 == null)
                continue;

            for (int i = 15; i < 23; i++) // iterate over categories in bio
            {
                Cell cell = row3.getCell(i, Row.CREATE_NULL_AS_BLANK);
                if (cell == null)
                    continue;

                String content = formatter.formatCellValue(cell);

                if (i == 15 & content.equals("1")) // check if post belong to definite category
                {
                    Row dest_row = sport.createRow(S_number);
                    copyRow(bios, sport, row3, dest_row, styleMap);
                    S_number += 1;
                } else if (i == 16 & content.equals("1")) // check if post belong to definite category
                {

                    Row dest_row = tech.createRow(T_number);

                    copyRow(bios, tech, row3, dest_row, styleMap);

                    T_number += 1;

                } else if (i == 17 & content.equals("1")) // check if post belong to definite category
                {

                    Row dest_row = business.createRow(B_number);

                    copyRow(bios, business, row3, dest_row, styleMap);

                    B_number += 1;

                } else if (i == 18 & content.equals("1")) // check if post belong to definite category
                {
                    Row dest_row = media.createRow(M_number);
                    copyRow(bios, media, row3, dest_row, styleMap);
                    M_number += 1;
                } else if (i == 19 & content.equals("1")) // check if post belong to definite category
                {
                    Row dest_row = youth.createRow(Y_number);
                    copyRow(bios, youth, row3, dest_row, styleMap);
                    Y_number += 1;
                } else if (i == 20 & content.equals("1")) // check if post belong to definite category
                {
                    Row dest_row = ceo.createRow(C_number);
                    copyRow(bios, ceo, row3, dest_row, styleMap);
                    C_number += 1;
                }
                if (i == 21 & content.equals("1")) // check if post belong to definite category
                {
                    Row dest_row = music.createRow(Mu_number);
                    copyRow(bios, music, row3, dest_row, styleMap);
                    Mu_number += 1;
                }
                if (i == 22 & content.equals("1")) // check if post belong to definite category
                {
                    Row dest_row = bloggers.createRow(Bl_number);
                    copyRow(bios, bloggers, row3, dest_row, styleMap);
                    Bl_number += 1;
                }
            }
        }

        copySheet(bios, cbios);
        copySheet(sport, csport);
        copySheet(tech, ctech); // make this and output and category detection
        copySheet(business, cbusiness);
        copySheet(media, cmedia);
        copySheet(youth, cyouth);
        copySheet(ceo, cceo);
        copySheet(music, cmusic);
        copySheet(bloggers, cbloggers);

        // Write the output to a file
        String file2 = "bios_all.xlsx";

        FileOutputStream out = new FileOutputStream(file2);
        wb_main.write(out);
        out.close();

        wb_main.close();

        // Write the output to a file sport
        String file3 = "sport.xlsx";

        FileOutputStream out1 = new FileOutputStream(file3);
        wb_sport.write(out1);
        out1.close();

        wb_sport.close();

        // reading file from desktop
        File inputFile = new File("sport.xlsx");
        // writing excel data to csv
        File outputFile = new File("sport.csv");
        xlsx(inputFile, outputFile);

        // Write the output to a file tech
        String file4 = "tech.xlsx";

        FileOutputStream out2 = new FileOutputStream(file4);
        wb_tech.write(out2);
        out2.close();

        wb_tech.close();

        // reading file from desktop
        File inputFile2 = new File("tech.xlsx");
        // writing excel data to csv
        File outputFile2 = new File("tech.csv");
        xlsx(inputFile2, outputFile2);

        // Write the output to a file business
        String file5 = "business.xlsx";

        FileOutputStream out3 = new FileOutputStream(file5);
        wb_business.write(out3);
        out3.close();

        wb_business.close();

        // reading file from desktop
        File inputFile3 = new File("business.xlsx");
        // writing excel data to csv
        File outputFile3 = new File("tech.csv");
        xlsx(inputFile3, outputFile3);

        // Write the output to a file media
        String file6 = "media.xlsx";

        FileOutputStream out4 = new FileOutputStream(file6);
        wb_media.write(out4);
        out4.close();

        wb_media.close();

        // reading file from desktop
        File inputFile4 = new File("media.xlsx");
        // writing excel data to csv
        File outputFile4 = new File("media.csv");
        xlsx(inputFile4, outputFile4);
    }

    static List<FormulaInfo> formulaInfoList = new ArrayList<FormulaInfo>();

    public static void refreshFormula(Workbook workbook) {
        for (FormulaInfo formulaInfo : formulaInfoList) {
            workbook.getSheet(formulaInfo.getSheetName()).getRow(formulaInfo.getRowIndex())
                    .getCell(formulaInfo.getCellIndex()).setCellFormula(formulaInfo.getFormula());
        }
        formulaInfoList.removeAll(formulaInfoList);
    }

    private static void copySheet(Sheet source, Sheet destination) {
        copySheet(source, destination, true);
    }

    /**
     * @param destination
     *                    the sheet to create from the copy.
     * @param the
     *                    sheet to copy.
     * @param copyStyle
     *                    true copy the style.
     */
    public static void copySheet(Sheet source, Sheet destination, boolean copyStyle) {
        int maxColumnNum = 0;
        List<CellStyle> styleMap = (copyStyle) ? new ArrayList<>() : null;

        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            Row srcRow = source.getRow(i);
            Row destRow = destination.createRow(i);
            if (srcRow != null) {
                copyRow(source, destination, srcRow, destRow, styleMap);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }
    }

    private static void copyRow(Sheet srcSheet, Sheet destSheet, Row srcRow, Row destRow,
            List<CellStyle> styleMap) {
        // manage a list of merged zone in order to not insert two times a
        // merged zone
        Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) {
            destRow.setHeight(srcRow.getHeight());
        }
        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }
        for (; j <= srcRow.getLastCellNum(); j++) {
            Cell oldCell = srcRow.getCell(j);
            Cell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                copyCell(oldCell, newCell, (List<CellStyle>) styleMap);

                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                        (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                    if (isNewMergedRegion(wrapper, mergedRegions)) {
                        mergedRegions.add(wrapper);
                        destSheet.addMergedRegion(wrapper.range);
                    }
                }
            }
        }

    }

    /**
     * 
     * @param sheet
     *                the sheet containing the data.
     * @param rowNum
     *                the num of the row to copy.
     * @param cellNum
     *                the num of the cell to copy.
     * @return the CellRangeAddress created.
     */
    public static CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    /**
     * Check that the merged region has been created in the destination sheet.
     * 
     * @param newMergedRegion
     *                        the merged region to copy or not in the destination
     *                        sheet.
     * @param mergedRegions
     *                        the list containing all the merged region.
     * @return true if the merged region is already in the list or not.
     */
    private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
            Set<CellRangeAddressWrapper> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }

    /**
     * @param oldCell
     * @param newCell
     * @param styleMap
     */
    private static void copyCell(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        if (styleList != null) {
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
                newCell.setCellStyle(oldCell.getCellStyle());
            } else {
                DataFormat newDataFormat = newCell.getSheet().getWorkbook().createDataFormat();
            }
        }
        switch (oldCell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_BLANK:
                newCell.setCellType(Cell.CELL_TYPE_BLANK);
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                try {
                    newCell.setCellFormula(oldCell.getCellFormula());
                } catch (Exception e) {

                }
                formulaInfoList.add(new FormulaInfo(oldCell.getSheet().getSheetName(), oldCell.getRowIndex(), oldCell
                        .getColumnIndex(), oldCell.getCellFormula()));
                break;
            default:
                break;
        }
    }

    private static CellStyle getSameCellStyle(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        CellStyle styleToFind = oldCell.getCellStyle();
        CellStyle currentCellStyle = null;
        CellStyle returnCellStyle = null;
        Iterator<CellStyle> iterator = styleList.iterator();
        Font oldFont = null;
        Font newFont = null;
        while (iterator.hasNext() && returnCellStyle == null) {
            currentCellStyle = iterator.next();

            if (currentCellStyle.getAlignment() != styleToFind.getAlignment()) {
                continue;
            }
            if (currentCellStyle.getHidden() != styleToFind.getHidden()) {
                continue;
            }
            if (currentCellStyle.getLocked() != styleToFind.getLocked()) {
                continue;
            }
            if (currentCellStyle.getWrapText() != styleToFind.getWrapText()) {
                continue;
            }
            if (currentCellStyle.getBorderBottom() != styleToFind.getBorderBottom()) {
                continue;
            }
            if (currentCellStyle.getBorderLeft() != styleToFind.getBorderLeft()) {
                continue;
            }
            if (currentCellStyle.getBorderRight() != styleToFind.getBorderRight()) {
                continue;
            }
            if (currentCellStyle.getBorderTop() != styleToFind.getBorderTop()) {
                continue;
            }
            if (currentCellStyle.getBottomBorderColor() != styleToFind.getBottomBorderColor()) {
                continue;
            }
            if (currentCellStyle.getFillBackgroundColor() != styleToFind.getFillBackgroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillForegroundColor() != styleToFind.getFillForegroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillPattern() != styleToFind.getFillPattern()) {
                continue;
            }
            if (currentCellStyle.getIndention() != styleToFind.getIndention()) {
                continue;
            }
            if (currentCellStyle.getLeftBorderColor() != styleToFind.getLeftBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRightBorderColor() != styleToFind.getRightBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRotation() != styleToFind.getRotation()) {
                continue;
            }
            if (currentCellStyle.getTopBorderColor() != styleToFind.getTopBorderColor()) {
                continue;
            }
            if (currentCellStyle.getVerticalAlignment() != styleToFind.getVerticalAlignment()) {
                continue;
            }

            oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
            newFont = newCell.getSheet().getWorkbook().getFontAt(currentCellStyle.getFontIndex());

            if (newFont.getBoldweight() != oldFont.getBoldweight()) {
                continue;
            }
            if (newFont.getColor() != oldFont.getColor()) {
                continue;
            }
            if (newFont.getFontHeight() != oldFont.getFontHeight()) {
                continue;
            }
            if (newFont.getFontName() != oldFont.getFontName()) {
                continue;
            }
            if (newFont.getItalic() != oldFont.getItalic()) {
                continue;
            }
            if (newFont.getStrikeout() != oldFont.getStrikeout()) {
                continue;
            }
            if (newFont.getTypeOffset() != oldFont.getTypeOffset()) {
                continue;
            }
            if (newFont.getUnderline() != oldFont.getUnderline()) {
                continue;
            }
            if (newFont.getCharSet() != oldFont.getCharSet()) {
                continue;
            }
            if (oldCell.getCellStyle().getDataFormatString().equals(currentCellStyle.getDataFormatString())) {
                continue;
            }

            returnCellStyle = currentCellStyle;
        }
        return returnCellStyle;
    }

    private static void copySheetSettings(Sheet newSheet, Sheet sheetToCopy) {

        newSheet.setAutobreaks(sheetToCopy.getAutobreaks());
        newSheet.setDefaultColumnWidth(sheetToCopy.getDefaultColumnWidth());
        newSheet.setDefaultRowHeight(sheetToCopy.getDefaultRowHeight());
        newSheet.setDefaultRowHeightInPoints(sheetToCopy.getDefaultRowHeightInPoints());
        newSheet.setDisplayGuts(sheetToCopy.getDisplayGuts());
        newSheet.setFitToPage(sheetToCopy.getFitToPage());

        newSheet.setForceFormulaRecalculation(sheetToCopy.getForceFormulaRecalculation());

        PrintSetup sheetToCopyPrintSetup = sheetToCopy.getPrintSetup();
        PrintSetup newSheetPrintSetup = newSheet.getPrintSetup();

        newSheetPrintSetup.setPaperSize(sheetToCopyPrintSetup.getPaperSize());
        newSheetPrintSetup.setScale(sheetToCopyPrintSetup.getScale());
        newSheetPrintSetup.setPageStart(sheetToCopyPrintSetup.getPageStart());
        newSheetPrintSetup.setFitWidth(sheetToCopyPrintSetup.getFitWidth());
        newSheetPrintSetup.setFitHeight(sheetToCopyPrintSetup.getFitHeight());
        newSheetPrintSetup.setLeftToRight(sheetToCopyPrintSetup.getLeftToRight());
        newSheetPrintSetup.setLandscape(sheetToCopyPrintSetup.getLandscape());
        newSheetPrintSetup.setValidSettings(sheetToCopyPrintSetup.getValidSettings());
        newSheetPrintSetup.setNoColor(sheetToCopyPrintSetup.getNoColor());
        newSheetPrintSetup.setDraft(sheetToCopyPrintSetup.getDraft());
        newSheetPrintSetup.setNotes(sheetToCopyPrintSetup.getNotes());
        newSheetPrintSetup.setNoOrientation(sheetToCopyPrintSetup.getNoOrientation());
        newSheetPrintSetup.setUsePage(sheetToCopyPrintSetup.getUsePage());
        newSheetPrintSetup.setHResolution(sheetToCopyPrintSetup.getHResolution());
        newSheetPrintSetup.setVResolution(sheetToCopyPrintSetup.getVResolution());
        newSheetPrintSetup.setHeaderMargin(sheetToCopyPrintSetup.getHeaderMargin());
        newSheetPrintSetup.setFooterMargin(sheetToCopyPrintSetup.getFooterMargin());
        newSheetPrintSetup.setCopies(sheetToCopyPrintSetup.getCopies());

        Header sheetToCopyHeader = sheetToCopy.getHeader();
        Header newSheetHeader = newSheet.getHeader();
        newSheetHeader.setCenter(sheetToCopyHeader.getCenter());
        newSheetHeader.setLeft(sheetToCopyHeader.getLeft());
        newSheetHeader.setRight(sheetToCopyHeader.getRight());

        Footer sheetToCopyFooter = sheetToCopy.getFooter();
        Footer newSheetFooter = newSheet.getFooter();
        newSheetFooter.setCenter(sheetToCopyFooter.getCenter());
        newSheetFooter.setLeft(sheetToCopyFooter.getLeft());
        newSheetFooter.setRight(sheetToCopyFooter.getRight());

        newSheet.setHorizontallyCenter(sheetToCopy.getHorizontallyCenter());
        newSheet.setMargin(Sheet.LeftMargin, sheetToCopy.getMargin(Sheet.LeftMargin));
        newSheet.setMargin(Sheet.RightMargin, sheetToCopy.getMargin(Sheet.RightMargin));
        newSheet.setMargin(Sheet.TopMargin, sheetToCopy.getMargin(Sheet.TopMargin));
        newSheet.setMargin(Sheet.BottomMargin, sheetToCopy.getMargin(Sheet.BottomMargin));

        newSheet.setPrintGridlines(sheetToCopy.isPrintGridlines());
        newSheet.setRowSumsBelow(sheetToCopy.getRowSumsBelow());
        newSheet.setRowSumsRight(sheetToCopy.getRowSumsRight());
        newSheet.setVerticallyCenter(sheetToCopy.getVerticallyCenter());
        newSheet.setDisplayFormulas(sheetToCopy.isDisplayFormulas());
        newSheet.setDisplayGridlines(sheetToCopy.isDisplayGridlines());
        newSheet.setDisplayRowColHeadings(sheetToCopy.isDisplayRowColHeadings());
        newSheet.setDisplayZeros(sheetToCopy.isDisplayZeros());
        newSheet.setPrintGridlines(sheetToCopy.isPrintGridlines());
        newSheet.setRightToLeft(sheetToCopy.isRightToLeft());
        newSheet.setZoom(1, 1);
        copyPrintTitle(newSheet, sheetToCopy);
    }

    private static void copyPrintTitle(Sheet newSheet, Sheet sheetToCopy) {
        int nbNames = sheetToCopy.getWorkbook().getNumberOfNames();
        Name name = null;
        String formula = null;

        String part1S = null;
        String part2S = null;
        String formS = null;
        String formF = null;
        String part1F = null;
        String part2F = null;
        int rowB = -1;
        int rowE = -1;
        int colB = -1;
        int colE = -1;

        for (int i = 0; i < nbNames; i++) {
            name = sheetToCopy.getWorkbook().getNameAt(i);
            if (name.getSheetIndex() == sheetToCopy.getWorkbook().getSheetIndex(sheetToCopy)) {
                if (name.getNameName().equals("Print_Titles")
                        || name.getNameName().equals(NameRecord.BUILTIN_PRINT_TITLE)) {
                    formula = name.getRefersToFormula();
                    int indexComma = formula.indexOf(",");
                    if (indexComma == -1) {
                        indexComma = formula.indexOf(";");
                    }
                    String firstPart = null;
                    ;
                    String secondPart = null;
                    if (indexComma == -1) {
                        firstPart = formula;
                    } else {
                        firstPart = formula.substring(0, indexComma);
                        secondPart = formula.substring(indexComma + 1);
                    }

                    formF = firstPart.substring(firstPart.indexOf("!") + 1);
                    part1F = formF.substring(0, formF.indexOf(":"));
                    part2F = formF.substring(formF.indexOf(":") + 1);

                    if (secondPart != null) {
                        formS = secondPart.substring(secondPart.indexOf("!") + 1);
                        part1S = formS.substring(0, formS.indexOf(":"));
                        part2S = formS.substring(formS.indexOf(":") + 1);
                    }

                    rowB = -1;
                    rowE = -1;
                    colB = -1;
                    colE = -1;
                    String rowBs, rowEs, colBs, colEs;
                    if (part1F.lastIndexOf("$") != part1F.indexOf("$")) {
                        rowBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        rowEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        rowB = Integer.parseInt(rowBs);
                        rowE = Integer.parseInt(rowEs);
                        if (secondPart != null) {
                            colBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            colEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                            colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);
                        }
                    } else {
                        colBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        colEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                        colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);

                        if (secondPart != null) {
                            rowBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            rowEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            rowB = Integer.parseInt(rowBs);
                            rowE = Integer.parseInt(rowEs);
                        }
                    }
                    newSheet.setRepeatingRows(CellRangeAddress.valueOf("rowB-1:rowE-1"));
                    newSheet.setRepeatingColumns(CellRangeAddress.valueOf("colB:colE"));
                }
            }
        }
    }

    private static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();

        short borderColor = IndexedColors.GREY_50_PERCENT.getIndex();

        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 48);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font sportFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(sportFont);
        styles.put("sport", style);

        Font techFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(techFont);
        styles.put("tech", style);

        Font businessFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(businessFont);
        styles.put("business", style);

        Font mediaFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIME.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(mediaFont);
        styles.put("media", style);

        Font youthFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LAVENDER.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(youthFont);
        styles.put("youth", style);

        Font ceoFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(ceoFont);
        styles.put("ceo", style);

        Font musicFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(musicFont);
        styles.put("music", style);

        Font bloggerFont = wb.createFont();

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(bloggerFont);
        styles.put("blogger", style);

        return styles;
    }

    public static void xlsx(File inputFile, File outputFile) {
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();

        try {
            FileOutputStream fos = new FileOutputStream(outputFile);
            // Get the workbook object for XLSX file
            XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(inputFile));
            // Get first sheet from the workbook
            Sheet sheet = wBook.getSheetAt(0);
            Row row;
            Cell cell;
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BOOLEAN:
                            data.append(cell.getBooleanCellValue() + ",");

                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            data.append(cell.getNumericCellValue() + ",");

                            break;
                        case Cell.CELL_TYPE_STRING:
                            data.append(cell.getStringCellValue() + ",");
                            break;

                        case Cell.CELL_TYPE_BLANK:
                            data.append("" + ",");
                            break;
                        default:
                            data.append(cell + ",");
                    }

                }
                data.append("\r\n");
            }

            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }
}
