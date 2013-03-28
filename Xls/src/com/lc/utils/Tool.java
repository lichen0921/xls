package com.lc.utils;

import java.io.File;
import java.io.FileFilter;
import java.util.ArrayList;
import java.util.List;

import android.os.Environment;

public class Tool {

	public static String getSDPath() {
		if (Environment.getExternalStorageState().equals(
				android.os.Environment.MEDIA_MOUNTED)) {
			return Environment.getExternalStorageDirectory().toString();
		}
		return null;
	}

	public static String getExcelPath() {
		String path = "/POI/MYDB/SUMMARY/";
		String sdPath = getSDPath();
		if (sdPath != null) {
			return sdPath + path;
		}
		return null;
	}

	public static List<File> getFiles(File file) {
		List<File> listFile = new ArrayList<File>();
		if (file.isDirectory()) {
			File[] files = file.listFiles(new FileFilter() {
				@Override
				public boolean accept(File pathname) {
					if (pathname.getName().toLowerCase()
							.substring(pathname.getName().lastIndexOf("."))
							.equals(".xls")) {
						return true;
					} else {
						return false;
					}
				}
			});

			for (int i = 0; i < files.length; i++) {
				File f = files[i];
				listFile.add(f);
			}
			return listFile;
		}
		return null;
	}
}
