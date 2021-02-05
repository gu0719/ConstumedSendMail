package util;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelHandleUtil {

    public static List<Map<String, String>> getExcelInfo(String fileName) {
        List<Map<String, String>> result = new ArrayList<>();

        //读取excel
        try {
            int realRows = getVaildRows(fileName);
            System.out.println("实际行数：" + realRows);
            if (realRows == 0) {
                return new ArrayList<>();
            }
            // 读取Excel文件内容
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(new File(fileName)));
            //读取xlsx中的sheet 这里只支持一页
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
            if (xssfSheet == null) {
                return new ArrayList<>();
            }
            //读取sheet的第一行作为表头
            XSSFRow xssfRow = xssfSheet.getRow(0);
            int cellCount = xssfRow.getLastCellNum();
            System.out.println("实际列数：" + cellCount);
            List<String> fieldList = new ArrayList<>();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                //前后去空格的String
                fieldList.add(xssfRow.getCell(cellNum).getStringCellValue().trim());
            }
            //装配数据key-value
            for (int rowNum = 1; rowNum <= realRows; rowNum++) {
                if (xssfSheet.getRow(rowNum) == null) {
                    continue;
                }
                XSSFRow otherRow = xssfSheet.getRow(rowNum);
                if (otherRow.getLastCellNum() < 0) {
                    continue;
                }
                //遍历每一行
                Map<String, String> map = new HashMap<>();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    if (otherRow.getCell(cellNum) != null) {
                        otherRow.getCell(cellNum).setCellType(Cell.CELL_TYPE_STRING);
                        map.put(fieldList.get(cellNum), otherRow.getCell(cellNum).getStringCellValue().trim());

                    }
                    //按照字段数进行装配
                }
                result.add(map);
            }
            System.out.println(result);
            System.out.println("应发邮件数：" + result.size());
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        return result;
    }

    public static int getVaildRows(String path) throws IOException, InvalidFormatException {
        FileInputStream excelFileInputStream = new FileInputStream(path);
        Workbook wb = WorkbookFactory.create(excelFileInputStream);
        Sheet sheet = wb.getSheetAt(0);
        CellReference cellReference = new CellReference("A4");
        boolean flag = false;
        for (int i = cellReference.getRow(); i <= sheet.getLastRowNum(); ) {
            Row r = sheet.getRow(i);
            if (r == null) {
                // 如果是空行（即没有任何数据、格式），直接把它以下的数据往上移动
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
                continue;
            }
            flag = false;
            for (Cell c : r) {
                if (c.getCellType() != Cell.CELL_TYPE_BLANK) {
                    flag = true;
                    break;
                }
            }
            if (flag) {
                i++;
                continue;
            } else {//如果是空白行（即可能没有数据，但是有一定格式）
                if (i == sheet.getLastRowNum())//如果到了最后一行，直接将那一行remove掉
                    sheet.removeRow(r);
                else//如果还没到最后一行，则数据往上移一行
                    sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
            }
        }
        return sheet.getLastRowNum() + 1;
    }


}
