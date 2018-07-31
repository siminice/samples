package org.rsfa.samples;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.ColumnHelper;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class Extractor {

    private static final int MAX_ROWS = 1000;
    private static final int MAX_COLS = 20;
    private static final short FONT_SIZE = 12;
    private static final short HEADER_SIZE = 14;

    private String inputFile;
    private String outputFile;

    private XSSFWorkbook wbin;
    private XSSFWorkbook wbout;

    private CellStyle idStyle;
    private CellStyle vStyle;
    private CellStyle hcolStyle;

    private String grid[][];
    private Set<String> ids = new HashSet<>();
    private int numRows = 0;
    private int numCols = 0;

    public Extractor() {
      grid = new String[MAX_ROWS][MAX_COLS];
    }

    public void setInputFile(String filename) {
        this.inputFile = filename;
    }

    public void setOutputFile(String filename) {
        this.outputFile = filename;
    }

    public void read() {
        try {

            FileInputStream excelFile = new FileInputStream(new File(inputFile));
            wbin = new XSSFWorkbook(excelFile);
            Sheet sheet = wbin.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    String v = currentCell.getStringCellValue();
                    if (v.contains("Row;Column")) continue;
                    String[] parts = v.split(";");
                    if (parts.length == 3) {
                        Integer r = Integer.parseInt(parts[0]);
                        Integer c = Integer.parseInt(parts[1]);
                        String id = parts[2];
                        if (r > numRows) numRows = r;
                        if (c > numCols) numCols = c;
                        if (r>0 && c>0) {
                            grid[r-1][c-1] = id;
                        }
                    } else {
                        System.err.println(String.format("Error parsing line: %s", v));
                    }
                }
            }
            System.out.println(String.format("Read %d rows x %d columns from %s",
                    numRows, numCols, inputFile));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public void extract() {
        wbout = new XSSFWorkbook();
        initStyles();
        XSSFSheet sheet = wbout.createSheet("Grid");
        Row header = sheet.createRow(0);
        header.createCell(0);
        header.createCell(1);
        for (int c = 0; c < numCols; c++) {
            Cell cellId = header.createCell(2*c+2);
            cellId.setCellStyle(hcolStyle);
            cellId.setCellValue(String.format("ID %d", c+1));
            sheet.setColumnWidth(2*c+2, 256*18);
            //sheet.autoSizeColumn(2*c+2);

            Cell cellValue = header.createCell(2*c+3);
            cellValue.setCellStyle(hcolStyle);
            cellValue.setCellValue(String.format("S %d", c+1));
            sheet.setColumnWidth(2*c+3, 256*6);
            //sheet.autoSizeColumn(2*c+3);
        }
        sheet.setAutobreaks(false);
        sheet.setRowBreak(50);
        sheet.setColumnBreak(24);
        sheet.setColumnHidden(1, true);
        for (int r = 0; r < numRows; r++) {
            Row row = sheet.createRow(r+1);
            Cell num = row.createCell(0);
            num.setCellValue(r+1);
            Cell hidden = row.createCell(1);

            for (int c = 0; c < numCols; c++) {
                if (grid[r][c] != null) {
                    Cell cellId = row.createCell(2*c+2);
                    cellId.setCellStyle(idStyle);
                    cellId.setCellValue(grid[r][c]);
                    ids.add(grid[r][c]);

                    Cell cellValue = row.createCell(2*c+3);
                    cellValue.setCellStyle(vStyle);
                }
            }
        }

        // List sheet
        List<String> uniqueIds = ids.stream().collect(Collectors.toList());
        Collections.sort(uniqueIds);
        System.out.println(String.format("Found %d unique sample ids.", uniqueIds.size()));

        XSSFSheet listSheet = wbout.createSheet("List");
        Row lheader = listSheet.createRow(0);
        Cell hnum = lheader.createCell(0);
        hnum.setCellStyle(hcolStyle);
        hnum.setCellValue("Anonymized ID");
        Cell sid = lheader.createCell(1);
        sid.setCellStyle(hcolStyle);
        sid.setCellValue("Sample ID");
        Cell score = lheader.createCell(2);
        score.setCellStyle(hcolStyle);
        score.setCellValue("Score");

        int numIds = 0;
        for (String id : uniqueIds) {
            Row rowi = listSheet.createRow(numIds+1);
            Cell c1 = rowi.createCell(0);
            c1.setCellValue(numIds+1);
            Cell c2 = rowi.createCell(1);
            c2.setCellValue(id);
            Cell c3 = rowi.createCell(2);
            numIds++;
        }
        listSheet.autoSizeColumn(0);
        listSheet.autoSizeColumn(1);
    }


    public void write() {
        try {
            FileOutputStream outputStream = new FileOutputStream(outputFile);
            wbout.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void initStyles() {
        Font fontNormal = wbout.createFont();
        fontNormal.setFontHeightInPoints(FONT_SIZE);
        Font fontHeader = wbout.createFont();
        fontHeader.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontHeader.setFontHeightInPoints(HEADER_SIZE);

        idStyle = wbout.createCellStyle();
        idStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        idStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        idStyle.setBorderTop(CellStyle.BORDER_THIN);
        idStyle.setBorderLeft(CellStyle.BORDER_THIN);
        idStyle.setBorderBottom(CellStyle.BORDER_THIN);

        vStyle = wbout.createCellStyle();
        vStyle.setBorderTop(CellStyle.BORDER_THIN);
        vStyle.setBorderRight(CellStyle.BORDER_THIN);
        vStyle.setBorderBottom(CellStyle.BORDER_THIN);

        hcolStyle = wbout.createCellStyle();
        hcolStyle.setFont(fontHeader);
        hcolStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        hcolStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        hcolStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        hcolStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        hcolStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        hcolStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        hcolStyle.setAlignment(CellStyle.ALIGN_CENTER);

        idStyle.setFont(fontNormal);
        vStyle.setFont(fontNormal);
        hcolStyle.setFont(fontHeader);
    }

}
