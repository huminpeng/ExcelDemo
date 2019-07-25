package com.leica.exceldemo;

import android.content.Context;
import android.util.Log;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

public class POIAndroid {
    private static final String TAG = "POIAndroid";

    private int insertStartRow = 0;

    private String sheetName = "Sheet1";

    private XSSFWorkbook getWorkBookHandle(String file) {
        XSSFWorkbook wb = null;
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
            wb = new XSSFWorkbook(fis);
        } catch (Exception e) {
            return null;
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return wb;
    }


    private XSSFRow createRow(XSSFSheet sheet, Integer rowIndex) {
        XSSFRow row = null;
        if (sheet.getRow(rowIndex) != null) {
            int lastRowNo = sheet.getLastRowNum();
            sheet.shiftRows(rowIndex, lastRowNo, 1);
        }
        row = sheet.createRow(rowIndex);
        return row;
    }

    private XSSFCell createCell(XSSFRow row, Record record) {
        XSSFCell cell = row.createCell((short) 0);
        cell.setCellValue(record.getName());
//        row.createCell(1).setCellValue(1.2);
//        row.createCell(2).setCellValue("This is a string cell");
        return cell;
    }

    private void saveExcel(String file, XSSFWorkbook wb) {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(file);
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public ArrayList<Record> readXLSXFileFromAssets(Context context, String fileName) {

        InputStream myInput = null;
        try {
            myInput = context.getAssets().open(fileName);

            XSSFWorkbook workbook = new XSSFWorkbook(myInput);
            XSSFSheet mySheet = workbook.getSheetAt(0);

            /** We now need something to iterate through the cells. **/
            Iterator<Row> rowIter = mySheet.rowIterator();

            ArrayList<Record> contactList = new ArrayList<>();

            while (rowIter.hasNext()) {
                Row myRow = (Row) rowIter.next();
                Iterator<Cell> cellIter = myRow.cellIterator();

                Record bean = new Record();

                while (cellIter.hasNext()) {

                    XSSFCell myCell = (XSSFCell) cellIter.next();
                    if (myCell.getColumnIndex() == 0) {

                        bean.setName(myCell.toString());
                    }

                    if (myCell.getColumnIndex() == 1) {

//                        bean.setMobile(myCell.toString());
                    }

                    Log.e("FileUtils", "Cell Value: " + myCell.toString() + " Index :" + myCell.getColumnIndex());

                }

                contactList.add(bean);
            }
            return contactList;

        } catch (IOException e) {
            e.printStackTrace();
            return null;
        } finally {
            if (myInput != null) {
                try {
                    myInput.close();
                } catch (IOException ioEx) {

                }
            }
        }
    }

    public ArrayList<Record> readXLSXFile(String file) {

        InputStream myInput = null;
        try {
            myInput = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(myInput);
            XSSFSheet mySheet = workbook.getSheetAt(0);

            /** We now need something to iterate through the cells. **/
            Iterator<Row> rowIter = mySheet.rowIterator();

            ArrayList<Record> contactList = new ArrayList<>();

            while (rowIter.hasNext()) {
                Row myRow = (Row) rowIter.next();
                Iterator<Cell> cellIter = myRow.cellIterator();

                Record bean = new Record();

                while (cellIter.hasNext()) {

                    XSSFCell myCell = (XSSFCell) cellIter.next();
                    if (myCell.getColumnIndex() == 0) {

                        bean.setName(myCell.toString());
                    }

                    if (myCell.getColumnIndex() == 1) {

//                        bean.setMobile(myCell.toString());
                    }

                    Log.e("FileUtils", "Cell Value: " + myCell.toString() + " Index :" + myCell.getColumnIndex());

                }

                contactList.add(bean);
            }
            return contactList;

        } catch (IOException e) {
            e.printStackTrace();
            return null;
        } finally {
            if (myInput != null) {
                try {
                    myInput.close();
                } catch (IOException ioEx) {

                }
            }
        }
    }

    public void create(String file, Record record) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet1 = workbook.createSheet(sheetName);
        XSSFRow row = createRow(sheet1, 0);
        createCell(row,record);
        saveExcel(file, workbook);
        return;
    }

    public void insertRows(String file, Record record) {
        XSSFWorkbook wb = getWorkBookHandle(file);
        XSSFSheet sheet1 = wb.getSheet(sheetName);
        XSSFRow row = createRow(sheet1, insertStartRow);
        createCell(row, record);
        saveExcel(file, wb);
    }

    public boolean merge(String file, int startRow, int endRow, int startCol, int endCol) {

        XSSFWorkbook workbook = getWorkBookHandle(file);

//        XSSFCellStyle style = workbook.createCellStyle();
//        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        XSSFSheet sheet1 = workbook.getSheet(sheetName);

        CellRangeAddress region = new CellRangeAddress(startRow, endRow, startCol, endCol);
        sheet1.addMergedRegion(region);
        saveExcel(file, workbook);
        return true;
    }

}
