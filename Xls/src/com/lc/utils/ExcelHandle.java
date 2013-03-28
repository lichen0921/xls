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
							x += rownum - 3;// -1 �Ǡ��˶������o�����B��
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
