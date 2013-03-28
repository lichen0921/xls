package com.lc.utils;

import java.io.File;
import java.util.List;

import android.util.Log;


import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcelHandle {

	private static final String TAG = "ExcelHandle";

	public static void read(File file) {
		try {
			Workbook book = Workbook.getWorkbook(file);
			Sheet sheet = book.getSheet(0);
			Cell cell1 = sheet.getCell(0, 5);
			String result = cell1.getContents();
			System.out.println(result);
			book.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static int sumTotal(List<File> sourceFiles, File targetFile) {
		try {
			if (targetFile.exists()) {
				sourceFiles.remove(targetFile);
				targetFile.delete();
			}

			Workbook book = null;
			Label label = null;
			int numberOfSheets = 0;
			String[] sheetNames = new String[] {};
			if (sourceFiles != null && sourceFiles.size() > 0) {
				File f0 = sourceFiles.get(0);
				book = Workbook.getWorkbook(f0);
				numberOfSheets = book.getNumberOfSheets();
				sheetNames = book.getSheetNames();

				WritableWorkbook writeBook = Workbook
						.createWorkbook(targetFile);
				for (int t = 0; t < numberOfSheets; t++) {
					WritableSheet writeSheet = writeBook.createSheet(
							sheetNames[t], t);
					int x = 0;
					for (int n = 0; n < sourceFiles.size(); n++) {
						book = Workbook.getWorkbook(sourceFiles.get(n));
						Sheet sheet = book.getSheet(t);
						sheet.getSettings().setProtected(true);
						sheet.getSettings().setPassword("123");
						if (sheet != null) {
							int columnum = sheet.getColumns();
							int rownum = sheet.getRows();
							if (n == 0) {
								for (int i = 0; i < 3; i++) {
									for (int j = 0; j < columnum; j++) {
										Cell cell = sheet.getCell(j, i);
										label = new Label(j, i,
												cell.getContents());
										writeSheet.addCell(label);
									}
								}
							}
							for (int i = 3; i < rownum; i++) {
								// int i = 1 means delete header from sheet
								for (int j = 0; j < columnum; j++) {
									Cell cell = sheet.getCell(j, i);
									label = new Label(j, i + x,
											cell.getContents());
									writeSheet.addCell(label);
								}
							}
							x += rownum - 3;// -1 是爲了多個表格無空行連接
							book.close();
						}
					}
				}
				writeBook.write();
				writeBook.close();
				return 1;
			}
		} catch (Exception e) {
			Log.e(TAG, "sumTotal error by lc");
			return -1;
		}
		return 0;
	}

	public static void write() {
		try {
			// 打开文件
			WritableWorkbook book = Workbook.createWorkbook(new File(
					"D:/temp/b.xls"));
			// 生成名为“第一页”的工作表，参数0表示这是第一页
			WritableSheet sheet = book.createSheet(" 第一页 ", 0);
			// 在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
			// 以及单元格内容为test
			Label label = new Label(0, 0, " test ");
			// 将定义好的单元格添加到工作表中
			sheet.addCell(label);
			/*
			 * 生成一个保存数字的单元格 必须使用Number的完整包路径，否则有语法歧义 单元格位置是第二列，第一行，值为789.123
			 */
			jxl.write.Number number = new jxl.write.Number(1, 0, 555.12541);
			sheet.addCell(number);
			// 写入数据并关闭文件
			book.write();
			book.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}

	public static void all() {
		try {
			Workbook book = Workbook.getWorkbook(new File("D:/temp/a.xls"));
			// 获得第一个工作表对象
			Sheet sheet = book.getSheet(0);
			// 得到第一列第一行的单元格
			int columnum = sheet.getColumns(); // 得到列数
			int rownum = sheet.getRows(); // 得到行数
			System.out.println(columnum);
			System.out.println(rownum);
			for (int i = 0; i < rownum; i++) // 循环进行读写
			{
				for (int j = 0; j < columnum; j++) {
					Cell cell1 = sheet.getCell(j, i);
					String result = cell1.getContents();
					System.out.print(result);
					System.out.print(" \t ");
				}
				System.out.println();
			}
			book.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
