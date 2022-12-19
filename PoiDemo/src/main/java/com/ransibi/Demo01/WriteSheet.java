package com.ransibi.Demo01;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * @description: 创建电子表格-sheet页
 * @author: rsb
 * @description: 2022-12-19-13-46
 * @description: 创建电子表格-sheet页
 * @Version: 1.0.0
 */
public class WriteSheet {
    public static void main(String[] args) throws Exception {
//        writeSheetMethod();
        readSheet();
    }

    public static void writeSheetMethod() throws Exception {
        //创建工作簿对象
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

        //创建工作表对象
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("员工信息");
        //创建行对象
        XSSFRow row;
        //构造数据
        Map<String, Object[]> employerInfoData = new TreeMap<String, Object[]>();
        employerInfoData.put("1", new Object[]{"编号", "名字", "描述"});
        employerInfoData.put("2", new Object[]{"1001", "张三", "开发"});
        employerInfoData.put("3", new Object[]{"1002", "李四", "开发"});
        employerInfoData.put("4", new Object[]{"1003", "王五", "运维"});
        employerInfoData.put("5", new Object[]{"1004", "小美", "测试"});
        employerInfoData.put("6", new Object[]{"1005", "李华", "产品"});

        //遍历数据，并写入工作表
        Set<String> keyId = employerInfoData.keySet();
        int rowId = 0;
        for (String key : keyId) {
            //根据工作表创建行对象
            row = xssfSheet.createRow(rowId++);
            Object[] objectArray = employerInfoData.get(key);
            int cellid = 0;
            for (Object obj : objectArray) {
                //根据行，创建单元格
                Cell cell = row.createCell(cellid++);
                //给单元格赋值
                cell.setCellValue((String) obj);
            }
        }
        FileOutputStream outputStream = new FileOutputStream(new File("员工信息表格.xlsx"));
        xssfWorkbook.write(outputStream);
        outputStream.close();
        System.out.println("数据写入excel成功!");
    }

    public static void readSheet() throws Exception {
        XSSFRow row;
        FileInputStream fileInputStream = new FileInputStream(new File("员工信息表格.xlsx"));
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
        Iterator<Row> rowIterator = xssfSheet.iterator();

        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + " \t\t ");
                        break;

                    case Cell.CELL_TYPE_STRING:
                        System.out.print(
                                cell.getStringCellValue() + " \t\t ");
                        break;
                    default:
                        break;
                }
            }
            System.out.println();
        }
        fileInputStream.close();

    }
}
