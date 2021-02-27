package com.romel.cathlabcalloff;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CathLabService {

    private SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

    //Read from lookup.
    private XSSFWorkbook workbookLookup;
    private XSSFSheet sheetLookup;
    private XSSFRow rowLookup;
    private XSSFCell cellLookup;

    //Read from previous report.
    private XSSFWorkbook workbookSource;
    private XSSFSheet sheetSource;
    private XSSFRow rowSource;
    private XSSFCell cellSource;

    //Create and write to target report.
    private XSSFWorkbook workbookTarget = new XSSFWorkbook();
    private XSSFSheet sheetTarget = workbookTarget.createSheet("Report");
    private XSSFRow rowTarget;


    //Target report cell styles.
    private DataFormat dataFormatTarget = workbookTarget.createDataFormat();
    private CellStyle dateStyle = workbookTarget.createCellStyle();
    private CellStyle decimalStyle = workbookTarget.createCellStyle();
    private CellStyle boldStyle = workbookTarget.createCellStyle();
    private CellStyle leftAlignStyle = workbookTarget.createCellStyle();
    
    Font fontBold = workbookTarget.createFont();
    
    public void run() {
        try {
            workbookLookup = new XSSFWorkbook(new FileInputStream("./lookup.xlsx"));
            sheetLookup = workbookLookup.getSheetAt(0);
            workbookSource = new XSSFWorkbook(new FileInputStream("./source.xlsx"));
            sheetSource = workbookSource.getSheetAt(0);

            dateStyle.setDataFormat(dataFormatTarget.getFormat("dd/MM/yyyy"));
            dateStyle.setAlignment(CellStyle.ALIGN_LEFT);
            decimalStyle.setDataFormat(dataFormatTarget.getFormat("000000"));

            fontBold.setBold(true);
            boldStyle.setFont(fontBold);
            leftAlignStyle.setAlignment(CellStyle.ALIGN_LEFT);

            //Write target title and date.
            int rowCountTarget = 0;
            rowTarget = sheetTarget.createRow(rowCountTarget);
            rowCountTarget ++;
            rowTarget.createCell(0).setCellValue("CATH LAB CALL OFF BALANCE");
            rowTarget.getCell(0).setCellStyle(boldStyle);
            rowTarget = sheetTarget.createRow(rowCountTarget);
            rowCountTarget ++;
            rowTarget.createCell(0).setCellStyle(dateStyle);
            rowTarget.getCell(0).setCellValue(new Date());

            //Create target column headers.
            rowTarget = sheetTarget.createRow(rowCountTarget);
            
            rowTarget.createCell(0).setCellValue("ORDER NUMBER");
            rowTarget.getCell(0).setCellStyle(boldStyle);
            rowTarget.createCell(1).setCellValue("SUPPLIER");
            rowTarget.getCell(1).setCellStyle(boldStyle);
            rowTarget.createCell(2).setCellValue("PART NUMBER");
            rowTarget.getCell(2).setCellStyle(boldStyle);
            rowTarget.createCell(3).setCellValue("DESCRIPTION");
            rowTarget.getCell(3).setCellStyle(boldStyle);
            rowTarget.createCell(4).setCellValue("ON ORDER");
            rowTarget.getCell(4).setCellStyle(boldStyle);
            rowTarget.createCell(5).setCellValue("OUTSTANDING");
            rowTarget.getCell(5).setCellStyle(boldStyle);
            rowTarget.createCell(6).setCellValue("COMMENTS");
            rowTarget.getCell(6).setCellStyle(boldStyle);
            rowTarget.createCell(7).setCellValue("NEW PO RAISED");
            rowTarget.getCell(7).setCellStyle(boldStyle);
            
            //Iterate through source workbook while generating values for target.
            //Iterate thru rows.
            boolean addNewPO = false;
            for(int i = 3; i <= sheetSource.getLastRowNum(); i ++) {
                rowCountTarget ++;
                rowSource = sheetSource.getRow(i);
                rowTarget = sheetTarget.createRow(rowCountTarget);
                int[] outstanding = null;
                
                //Iterate thru columns/cells.
                for(int j = 0; j < 8; j ++) {
                    cellSource = rowSource.getCell(j);
                    String newPO;
                    if(cellSource != null) {
                        switch (cellSource.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                rowTarget.createCell(j).setCellValue(cellSource.getStringCellValue());
                                //If current cell points to the outstanding column, do:
                                if(j == 0) {
                                    outstanding = getOutstanding(cellSource.getStringCellValue().trim());
                                }
                                else if(j == 4) {
                                    rowTarget.createCell(j).setCellValue(outstanding[0]);
                                }
                                else if(j == 5) {
                                    rowTarget.createCell(j).setCellValue(outstanding[1]);
                                }
                                else if(j == 7) {
                                    if(rowSource.getCell(7) != null || rowSource.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) {
                                        newPO = rowSource.getCell(7).getStringCellValue().trim();
                                        rowTarget.getCell(7).setCellValue("");
                                        addNewPO = true;
                                    }
                                }
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                if(DateUtil.isCellDateFormatted(cellSource)) {
                                    rowTarget.createCell(j).setCellValue(cellSource.getDateCellValue());
                                }
                                if(j == 4) {
                                    rowTarget.createCell(j).setCellValue(outstanding[0]);
                                }
                                else if(j == 5) {
                                    rowTarget.createCell(j).setCellValue(outstanding[1]);
                                }
                                else {
                                    rowTarget.createCell(j).setCellValue(cellSource.getNumericCellValue());
                                }
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                if(j == 4) {
                                    rowTarget.createCell(j).setCellValue(outstanding[0]);
                                }
                                else if(j == 5) {
                                    rowTarget.createCell(j).setCellValue(outstanding[1]);
                                }
                                break;
                        }
                    }
                    else {
                        if(j == 4) {
                            rowTarget.createCell(j).setCellValue(outstanding[0]);
                        }
                        else if(j == 5) {
                            rowTarget.createCell(j).setCellValue(outstanding[1]);
                        }
                    }

                    if(j == 2) {
                        rowTarget.getCell(2).setCellStyle(leftAlignStyle);
                    }
                }

                if(addNewPO) {
                    rowCountTarget ++;
                    addNewPo(rowSource.getCell(7).getStringCellValue().trim(), rowCountTarget,
                            rowSource.getCell(1).getStringCellValue().trim(), rowSource.getCell(2).getStringCellValue().trim(),
                            rowSource.getCell(3).getStringCellValue().trim());
                        //rowTarget.getCell(7).setCellValue("");
                        //lastRowNum ++;
                        addNewPO = false;
                }
            }

            //Set each cell to auto-resize.
            sheetTarget.autoSizeColumn(0);
            sheetTarget.autoSizeColumn(1);
            sheetTarget.autoSizeColumn(2);
            sheetTarget.autoSizeColumn(3);
            sheetTarget.autoSizeColumn(4);
            sheetTarget.autoSizeColumn(5);
            sheetTarget.autoSizeColumn(6);
            sheetTarget.autoSizeColumn(7);

            //Save changes and close workbooks.
            workbookTarget.write(new FileOutputStream("./CathLab_CallOff_Report.xlsx"));
            workbookTarget.close();
            workbookLookup.close();
            workbookLookup.close();
        }
        catch(IOException ex) {
            ex.printStackTrace();
        }
        catch(Exception ex) {
            ex.printStackTrace();
        }

    }//public void run().

    /**
     * 
     * @param purchaseOrder
     * @return int -> returns the outstanding on order.
     */
    public int[] getOutstanding(String purchaseOrder) {
        int[] result = {0,0};

        try {
            for(int i = 0; i <= sheetLookup.getLastRowNum(); i ++) {
                rowLookup = sheetLookup.getRow(i);
                if(rowLookup != null) {
                    cellLookup = rowLookup.getCell(2);
                    if(cellLookup != null) {
                        if(cellLookup != null) {
                            if(purchaseOrder.trim().equalsIgnoreCase(cellLookup.getStringCellValue().trim())) {
                                result[0] = (int)rowLookup.getCell(20).getNumericCellValue();
                                result[1] = (int)rowLookup.getCell(21).getNumericCellValue();
                            }
                        }
                    }
                }
            }

            return result;
        }
        catch(Exception ex) {
            ex.printStackTrace();
            return result;
        }
    }

    /**
     * Adds new purchase order to target.
     */
    public void addNewPo(String purchaseOrder, int rowPosition, String supplier, String partNumber, String description) {
        int[] outstanding = getOutstanding(purchaseOrder.trim());

        rowTarget = sheetTarget.createRow(rowPosition);
        rowTarget.createCell(0).setCellValue(purchaseOrder);
        rowTarget.createCell(1).setCellValue(supplier);
        rowTarget.createCell(2).setCellValue(partNumber);
        rowTarget.createCell(3).setCellValue(description);
        rowTarget.createCell(4).setCellValue(outstanding[0]);
        rowTarget.createCell(5).setCellValue(outstanding[1]);
    }

    /**
     * Not implemented.
     */
    public boolean isPOUnique(String purchaseOrder) {
        int poCount = 0;
        boolean isUnique = false;

        try {
            for(int i = 0; i <= sheetLookup.getLastRowNum(); i ++) {
                rowLookup = sheetLookup.getRow(i);
                if(rowLookup != null) {
                    cellLookup = rowLookup.getCell(2);
                    if(cellLookup != null) {
                        if(cellLookup != null) {
                            if(purchaseOrder.trim().equalsIgnoreCase(cellLookup.getStringCellValue().trim())) {
                                poCount ++;
                            }
                        }
                    }
                }
            }

            if(poCount > 1) {
                isUnique = false;
            }
            else if(poCount == 1) {
                isUnique = true;
            }
            else {
                isUnique = false;
            }

            return isUnique;
        }
        catch(Exception ex) {
            ex.printStackTrace();
            return isUnique;
        }
    }
    
}