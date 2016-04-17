package com.imooc.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Created by zZ on 2016-4-17.
 */
public class PoiTest {


	public static void main(String[] args) throws Exception{
		String fileName = "d:/1_Temp/poi_test.xls";
		createExcel(fileName);
		readExcel(fileName);
	}

	/**
	 * 通过文件的绝对路径（file）来生成固定表格
	 * @param fileName
	 * @throws Exception
	 */
	private static void createExcel(String fileName) throws Exception {
		// 创建空文件
		File file = new File(fileName);
		file.createNewFile();
		// 创建工作簿（WorkBook）
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 创建Sheet
		HSSFSheet sheet = workbook.createSheet();
		// 创建表头
		String[] title = {"id", "name", "sex"};
		HSSFRow row = sheet.createRow(0);
		for (int i = 0; i < title.length; i++) {
			HSSFCell cell = row.createCell(i);
			cell.setCellValue(title[i]);
		}
		// 添加数据
		for (int i = 1; i < 10; i++) {
			row = sheet.createRow(i);
			HSSFCell cell = row.createCell(0);
			cell.setCellValue("id" + i);
			cell = row.createCell(1);
			cell.setCellValue("user" + i);
			cell = row.createCell(2);
			cell.setCellValue("男");
		}
		// 将数据写入文件
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		// close stream
		workbook.close();
		fos.close();
	}

	/**
	 * 通过excel文件的绝对路径（fileName）来读取内容，并将内容输出至Console
	 * @param fileName
	 * @throws Exception
	 */
	private static void readExcel(String fileName) throws Exception{
		// 获取文件
		File file = new File(fileName);
		// 获取工作簿WorkBook
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
		// 获取Sheet
		// HSSFSheet sheet = workbook.getSheet("sheet0");  // 根据名字获取
		HSSFSheet sheet = workbook.getSheetAt(0);
		// 获取数据并输出
		int firstRow = sheet.getFirstRowNum();  // 第一行行号
		int lastRow = sheet.getLastRowNum();  // 最后一行行号
		for (int i = firstRow; i < lastRow; i++) {
			HSSFRow row = sheet.getRow(i);
			int firstCol = row.getFirstCellNum();  // 第一列列号
			int lastCol = row.getLastCellNum();  // 最后一列列号
			for (int j = firstCol; j < lastCol; j++) {
				HSSFCell cell = row.getCell(j);
				String value = cell.getStringCellValue();
				System.out.print(value + "  ");
			}
			System.out.println();
		}
	}
}
