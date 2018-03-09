package com.njuits.its.utils;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * com.njuits.its.utils Toolkit.java
 * 
 * @author wangh 804310112@qq.com 2017-11-17 下午6:26:33
 */
public class Toolkit {

	public static void safeClose(FileOutputStream fos) {
		if (fos != null) {
			try {
				fos.flush();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			try {
				fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public static void safeClose(XSSFWorkbook workbook) {
		if (workbook != null)
			try {
				workbook.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
	}

	@SafeVarargs
	public static <T> void safeClose(T... args) {
		if (args != null)
			for (Object arg : args) {
				if (arg instanceof FileOutputStream)
					safeClose((FileOutputStream) arg);
				else if (arg instanceof XSSFWorkbook)
					safeClose((XSSFWorkbook) arg);
			}
	}
}
