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
				// ����--�պ�
				WritableFont wfNormalblack = new WritableFont(
						WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false);
				WritableCellFormat cfNormalblack = new WritableCellFormat(
						wfNormalblack);
				cfNormalblack.setAlignment(Alignment.CENTRE);
				cfNormalblack.setVerticalAlignment(VerticalAlignment.CENTRE);
				cfNormalblack.setBorder(Border.ALL, BorderLineStyle.THIN);
				// ����--����
				WritableFont wfNormalblue = new WritableFont(
						WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false,
						UnderlineStyle.NO_UNDERLINE, Colour.BLUE);
				WritableCellFormat cfNormalblue = new WritableCellFormat(
						wfNormalblue);
				cfNormalblue.setAlignment(Alignment.CENTRE);
				cfNormalblue.setVerticalAlignment(VerticalAlignment.CENTRE);
				cfNormalblue.setBorder(Border.ALL, BorderLineStyle.THIN);
				// ������Ԫ��
				WritableCellFormat cfUnlock = new WritableCellFormat(
						wfNormalblue);
				cfUnlock.setAlignment(Alignment.CENTRE);
				cfUnlock.setVerticalAlignment(VerticalAlignment.CENTRE);
				cfUnlock.setBorder(Border.ALL, BorderLineStyle.THIN);
				cfUnlock.setLocked(false);

				writeBook = Workbook.createWorkbook(targetFile);
				for (int t = 0; t < numberOfSheets; t++) {
					writeSheet = writeBook.createSheet(sheetNames[t], t);
					// ����excel
					writeSheet.getSettings().setProtected(true);
					// ��һ�� ����
					WritableFont mWritableFont_title = new WritableFont(
							WritableFont.TIMES, 10, WritableFont.BOLD, false);
					WritableCellFormat mWritableCellFormat_title = new WritableCellFormat(
							mWritableFont_title);
					mWritableCellFormat_title.setAlignment(Alignment.CENTRE);
					mWritableCellFormat_title
							.setVerticalAlignment(VerticalAlignment.CENTRE);
					mWritableCellFormat_title.setBorder(Border.ALL,
							BorderLineStyle.THIN);
					Label labelTitle = new Label(0, 0, "�ɼ�������ȿ��Ƽ�¼��",
							mWritableCellFormat_title);
					writeSheet.addCell(labelTitle);
					// �ڶ��� ����
					if (t == 0) {
						writeSheet.mergeCells(0, 0, 18, 0);
						writeSheet.mergeCells(0, 1, 5, 1);
						writeSheet.mergeCells(6, 1, 13, 1);
						writeSheet.mergeCells(14, 1, 18, 1);
						// Label label02 = new Label(0, 2,
						// "�ɼ�������",cfNormalblack);
						// writeSheet.addCell(label02);
						// Label label12 = new Label(1, 2,
						// "���������",cfNormalblack);
						// writeSheet.addCell(label12);
						// Label label22 = new Label(2, 2,
						// "����λ��",cfNormalblack);
						// writeSheet.addCell(label22);
						// Label label32 = new Label(3, 2,
						// "δ�仯����",cfNormalblack);
						// writeSheet.addCell(label32);
						// Label label42 = new Label(4, 2,
						// "�仯����",cfNormalblack);
						// writeSheet.addCell(label42);
						// Label label52 = new Label(5, 2,
						// "��������",cfNormalblack);
						// writeSheet.addCell(label52);
						// Label label62 = new Label(6, 2,
						// "��·��Ϣ",cfNormalblack);
						// writeSheet.addCell(label62);
						// Label label72 = new Label(7, 2,
						// "��ֹ��Ϣ",cfNormalblack);
						// writeSheet.addCell(label72);
						// Label label82 = new Label(8, 2,
						// "ɾ��������",cfNormalblack);
						// writeSheet.addCell(label82);
						// Label label92 = new Label(9, 2,
						// "POI����",cfNormalblack);
						// writeSheet.addCell(label92);
						// Label label102 = new Label(10, 2,
						// "��ʱͶ��",cfNormalblack);
						// writeSheet.addCell(label102);
						// Label label112 = new Label(11, 2,
						// "�������",cfNormalblack);
						// writeSheet.addCell(label112);
						// Label label122 = new Label(12, 2,
						// "���Ч��",cfNormalblack);
						// writeSheet.addCell(label122);
						// Label label132 = new Label(13, 2,
						// "POIЧ��",cfNormalblack);
						// writeSheet.addCell(label132);
						// Label label142 = new Label(14, 2,
						// "��·��ϢЧ��",cfNormalblack);
						// writeSheet.addCell(label142);
						// Label label152 = new Label(15, 2,
						// "POI�ܶ�",cfNormalblack);
						// writeSheet.addCell(label152);
						// Label label162 = new Label(16, 2,
						// "�仯��",cfNormalblack);
						// writeSheet.addCell(label162);
						// Label label172 = new Label(17, 2,
						// "������",cfNormalblack);
						// writeSheet.addCell(label172);
						// Label label182 = new Label(18, 2,
						// "ɾ����",cfNormalblack);
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
								for (int i = 1; i < 3; i++) {// �ڶ�������
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
																// �Ǡ��˶������o�����B��
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
							.addCell(new Label(0, rows, "��  ��", cfNormalblue));

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
			// ���ļ�
			WritableWorkbook book = Workbook.createWorkbook(new File(
					"D:/temp/b.xls"));
			// ������Ϊ����һҳ���Ĺ���������0��ʾ���ǵ�һҳ
			WritableSheet sheet = book.createSheet(" ��һҳ ", 0);
			// ��Label����Ĺ�������ָ����Ԫ��λ���ǵ�һ�е�һ��(0,0)
			// �Լ���Ԫ������Ϊtest
			Label label = new Label(0, 0, " test ");
			// ������õĵ�Ԫ����ӵ���������
			sheet.addCell(label);
			/*
			 * ����һ���������ֵĵ�Ԫ�� ����ʹ��Number��������·�����������﷨���� ��Ԫ��λ���ǵڶ��У���һ�У�ֵΪ789.123
			 */
			jxl.write.Number number = new jxl.write.Number(1, 0, 555.12541);
			sheet.addCell(number);
			// д�����ݲ��ر��ļ�
			book.write();
			book.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}

	public static void all() {
		try {
			Workbook book = Workbook.getWorkbook(new File("D:/temp/a.xls"));
			// ��õ�һ�����������
			Sheet sheet = book.getSheet(0);
			// �õ���һ�е�һ�еĵ�Ԫ��
			int columnum = sheet.getColumns(); // �õ�����
			int rownum = sheet.getRows(); // �õ�����
			System.out.println(columnum);
			System.out.println(rownum);
			for (int i = 0; i < rownum; i++) // ѭ�����ж�д
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
