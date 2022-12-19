package com.ransibi.Demo01;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @description: 创建工作簿
 * @author: rsb
 * @description: 2022-12-19-11-40
 * @description: 创建工作簿
 * @Version: 1.0.0
 */
public class CreateWorkBook {
    public static void main(String[] args) throws Exception {
        //创建工作簿
        //createWorkBookMethod();

        //打开工作簿
        openWorkBookMethod();
    }


    public static void createWorkBookMethod() throws Exception {
        //创建工作簿对象
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //创建输出流对象，并初始化
        FileOutputStream outputStream = new FileOutputStream(new File("测试.xlsx"));
        //使用文件输出对象写入操作工作簿
        xssfWorkbook.write(outputStream);
        //关闭输出流
        outputStream.close();
        System.out.println("创建工作簿成功!");
    }

    public static void openWorkBookMethod() throws Exception {
        File file = new File("测试.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        //获取XLSX文件的工作簿实例
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        //打开工作簿后，就可以对其执行读写操作。
        if (file.isFile() && file.exists()) {
            System.out.println("打开文件成功!");
        } else {
            System.out.println("打开文件失败!");
        }

    }

}
