package com.lc.utils;

import java.io.File;
import java.util.List;

import android.util.Log;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcelHandle {

	private static final String TAG = "ExcelHandle";
	private Workbook book;
	private Label label;
	private WritableWorkbook writeBook;
	private WritableSheet writeSheet;

	public int sumTotal(List<File> sourceFiles, File targetFile) {
		try {
			if (targetFile.exists()) {
				sourceFiles.remove(targetFile);
				targetFile.delete();
			}

			book = null;
			label = null;
			int numberOfSheets = 0;
			String[] sheetNames = new String[] {};
			if (sourceFiles != null && sourceFiles.size() > 0) {
				File f0 = sourceFiles.get(0);
				book = Workbook.getWorkbook(f0);
				numberOfSheets = book.getNumberOfSheets();
				sheetNames = book.getSheetNames();
				// 字体--普黑
				WritableFont wfNormalblack = new WritableFont(
						WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false);
				WritableCellFormat cfNormalblack = new WritableCellFormat(
						wfNormalblack);
				cfNormalblack.setAlignment(Alignment.CENTRE);
				cfNormalblack.setVerticalAlignment(VerticalAlignment.CENTRE);
				cfNormalblack.setBorder(Border.ALL, BorderLineStyle.THIN);
				// 字体--普蓝
				WritableFont wfNormalblue = new WritableFont(
						WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false,
						UnderlineStyle.NO_UNDERLINE, Colour.BLUE);
				WritableCellFormat cfNormalblue = new WritableCellFormat(
						wfNormalblue);
				cfNormalblue.setAlignment(Alignment.CENTRE);
				cfNormalblue.setVerticalAlignment(VerticalAlignment.CENTRE);
				cfNormalblue.setBorder(Border.ALL, BorderLineStyle.THIN);
				// 解锁单元格
				WritableCellFormat cfUnlock = new WritableCellFormat(
						wfNormalblue);
				cfUnlock.setAlignment(Alignment.CENTRE);
				cfUnlock.setVerticalAlignment(VerticalAlignment.CENTRE);
				cfUnlock.setBorder(Border.ALL, BorderLineStyle.THIN);
				cfUnlock.setLocked(false);

				writeBook = Workbook.createWorkbook(targetFile);
				for (int t = 0; t < numberOfSheets; t++) {
					writeSheet = writeBook.createSheet(sheetNames[t], t);
					// 锁定excel
					writeSheet.getSettings().setProtected(true);
					// 第一行 标题
					WritableFont mWritableFont_title = new WritableFont(
							WritableFont.TIMES, 10, WritableFont.BOLD, false);
					WritableCellFormat mWritableCellFormat_title = new WritableCellFormat(
							mWritableFont_title);
					mWritableCellFormat_title.setAlignment(Alignment.CENTRE);
					mWritableCellFormat_title
							.setVerticalAlignment(VerticalAlignment.CENTRE);
					mWritableCellFormat_title.setBorder(Border.ALL,
							BorderLineStyle.THIN);
					Label labelTitle = new Label(0, 0, "采集任务进度控制记录表",
							mWritableCellFormat_title);
					writeSheet.addCell(labelTitle);
					// 第二行 标题
					if (t == 0) {
						writeSheet.mergeCells(0, 0, 18, 0);
						writeSheet.mergeCells(0, 1, 5, 1);
						writeSheet.mergeCells(6, 1, 13, 1);
						writeSheet.mergeCells(14, 1, 18, 1);
						// Label label02 = new Label(0, 2,
						// "采集区域编号",cfNormalblack);
						// writeSheet.addCell(label02);
						// Label label12 = new Label(1, 2,
						// "任务接收人",cfNormalblack);
						// writeSheet.addCell(label12);
						// Label label22 = new Label(2, 2,
						// "区域位置",cfNormalblack);
						// writeSheet.addCell(label22);
						// Label label32 = new Label(3, 2,
						// "未变化点数",cfNormalblack);
						// writeSheet.addCell(label32);
						// Label label42 = new Label(4, 2,
						// "变化点数",cfNormalblack);
						// writeSheet.addCell(label42);
						// Label label52 = new Label(5, 2,
						// "新增点数",cfNormalblack);
						// writeSheet.addCell(label52);
						// Label label62 = new Label(6, 2,
						// "道路信息",cfNormalblack);
						// writeSheet.addCell(label62);
						// Label label72 = new Label(7, 2,
						// "禁止信息",cfNormalblack);
						// writeSheet.addCell(label72);
						// Label label82 = new Label(8, 2,
						// "删除点数量",cfNormalblack);
						// writeSheet.addCell(label82);
						// Label label92 = new Label(9, 2,
						// "POI总数",cfNormalblack);
						// writeSheet.addCell(label92);
						// Label label102 = new Label(10, 2,
						// "人时投入",cfNormalblack);
						// writeSheet.addCell(label102);
						// Label label112 = new Label(11, 2,
						// "单向里程",cfNormalblack);
						// writeSheet.addCell(label112);
						// Label label122 = new Label(12, 2,
						// "里程效率",cfNormalblack);
						// writeSheet.addCell(label122);
						// Label label132 = new Label(13, 2,
						// "POI效率",cfNormalblack);
						// writeSheet.addCell(label132);
						// Label label142 = new Label(14, 2,
						// "道路信息效率",cfNormalblack);
						// writeSheet.addCell(label142);
						// Label label152 = new Label(15, 2,
						// "POI密度",cfNormalblack);
						// writeSheet.addCell(label152);
						// Label label162 = new Label(16, 2,
						// "变化率",cfNormalblack);
						// writeSheet.addCell(label162);
						// Label label172 = new Label(17, 2,
						// "新增率",cfNormalblack);
						// writeSheet.addCell(label172);
						// Label label182 = new Label(18, 2,
						// "删除率",cfNormalblack);
						// writeSheet.addCell(label182);
					} else if (t == 1) {
						writeSheet.mergeCells(0, 0, 20, 0);
						writeSheet.mergeCells(0, 1, 4, 1);
						writeSheet.mergeCells(5, 1, 13, 1);
						writeSheet.mergeCells(14, 1, 20, 1);
					}
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
								for (int i = 1; i < 3; i++) {// 第二，三行
									for (int j = 0; j < columnum; j++) {
										Cell cell = sheet.getCell(j, i);
										label = new Label(j, i,
												cell.getContents(),
												cfNormalblack);
										writeSheet.addCell(label);
									}
								}
							}
							for (int j = 0; j < columnum; j++) {
								// int i = 3 means delete header from sheet
								for (int i = 3; i < rownum - (t == 0 ? 1 : 0); i++) {
									Cell cell = sheet.getCell(j, i);
									String contents = cell.getContents();
									if (isNumber(contents)) {
										Number number = new Number(j, i + x,
												Double.valueOf(contents),
												cfNormalblue);
										writeSheet.addCell(number);
									} else {
										label = new Label(j, i + x, contents,
												cfNormalblue);
										writeSheet.addCell(label);
									}
								}
							}
							x += rownum - 3 - (t == 0 ? 1 : 0);// -1
																// 是爲了多個表格無空行連接
							book.close();
						}
					}
					writeSheet = writeBook.getSheet(t);
					int rows = writeSheet.getRows();
					for (int j = 0; j < writeSheet.getColumns(); j++) {
						Double sumCol = 0.0;
						Boolean flag = false;
						for (int i = 0; i < rows; i++) {
							Cell cell = writeSheet.getCell(j, i);
							String contents = cell.getContents();
							if (isNumber(contents)) {
								sumCol += Double.valueOf(contents);
								flag = true;
							} else {
								sumCol += 0.0;
							}
						}
						if (flag) {
							writeSheet.addCell(new Number(j, rows, sumCol,
									cfNormalblue));
						} else {
							writeSheet.addCell(new Label(j, rows, "",
									cfNormalblue));
						}
					}
					writeSheet
							.addCell(new Label(0, rows, "合  计", cfNormalblue));

				}
				writeBook.write();
				writeBook.close();
				return 1;
			}
		} catch (Exception e) {
			Log.e(TAG, "sumTotal error by lc");
			e.printStackTrace();
			return -1;
		}
		return 0;
	}

	public static boolean isNumber(String numberStr) {
		String valid = "0123456789.";
		String temp = "";
		boolean flag = true;
		if (numberStr == null || numberStr.equals("")
				|| numberStr.equals("null")) {
			flag = false;
		} else {
			for (int i = 0; i < numberStr.length(); i++) {
				temp = "" + numberStr.substring(i, i + 1);
				if (valid.indexOf(temp) == -1) {
					flag = false;
				}
			}
		}
		return flag;
	}

	// private Double calculateAdd(Sheet sheet) {
	// int columnum = sheet.getColumns();
	// int rownum = sheet.getRows();
	// Double sum = 0.0;
	// for (int j = 0; j < columnum; j++) {
	// Double sumCol = 0.0;
	// for (int i = 0; i < rownum; i++) {
	// Cell cell = sheet.getCell(j, i);
	// String contents = cell.getContents();
	// if (isNumber(contents)) {
	// sumCol += Double.valueOf(contents);
	// }
	// }
	// Number number = new Number(j, i + x, sumCol, cfNormalblue);
	// writeSheet.addCell(number);
	// }
	// return null;
	// }

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
