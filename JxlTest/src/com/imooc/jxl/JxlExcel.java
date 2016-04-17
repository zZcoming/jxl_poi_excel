package com.imooc.jxl;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;
import java.io.IOException;

/**
 * Created by zZ on 2016-4-17.
 */
public class JxlExcel {

	public static void main(String[] args) throws Exception{
		// fileName:生成excel文件的路径
		String fileName = "d:/1_Temp/jxl_test.xls";
		createExcel(fileName);
		readExcel(fileName);
	}

	/**
	 * 通过文件的绝对路径（file）来生成固定表格
	 * @param fileName
	 * @throws Exception
	 */
	public static void createExcel(String fileName) throws Exception{
		//创建Excel文件
		File f = new File(fileName);
		f.createNewFile();
		// 创建WorkBook
		WritableWorkbook workbook = Workbook.createWorkbook(f);
		// 创建Sheet
		WritableSheet sheet = workbook.createSheet("sheet_test", 0);
		// 创建数据
		Label label = null;
		// 添加表头
		String[] title = {"id", "name", "sex"};
		for(int i=0;i<title.length;i++) {
			// 单元格根据行号和列号来定位
			label = new Label(i,0,title[i]);
			sheet.addCell(label);
		}
		// 添加数据
		// 从1开始，否则会覆盖表头
		for(int i=1;i<10;i++) {
			label = new Label(0,i,"id" + i);
			sheet.addCell(label);
			label = new Label(1,i,"user" + i);
			sheet.addCell(label);
			label = new Label(2,i,"男");
			sheet.addCell(label);
		}
		workbook.write();
		workbook.close();
	}

	/**
	 * 通过excel文件的绝对路径（fileName）来读取内容，并将内容输出至Console
	 * @param fileName
	 * @throws Exception
	 */
	public static void readExcel(String fileName) throws Exception{
		// 获取Excel文件
		File f = new File(fileName);
		// 获取WorkBook（读取用的是WorkBook）
		Workbook workbook = Workbook.getWorkbook(f);
		// 获取第一个Sheet页
		Sheet sheet = workbook.getSheet(0);
		// 获取数据
		for(int i=0;i<sheet.getRows();i++) {
			for(int j=0;j<sheet.getColumns();j++) {
				// getCell(int col, int row)  注意顺序
				Cell cell = sheet.getCell(j,i);
				System.out.print(cell.getContents() + "\t");
			}
			System.out.println();
		}
		workbook.close();
	}

}
