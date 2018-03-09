package com.njuits.its;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.record.cf.FontFormatting;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.njuits.its.utils.Toolkit;

/**
 * com.njuits.its ExcelUtils.java
 * 
 * @author wangh 804310112@qq.com 2017-11-17 下午5:42:54
 */
public class ExcelUtils {

	public static void main(String[] args) {
		excelWriter("E:/hos.xlsx");
	}
	public static void excelWriter(String path) {
		// 新建一输出流并把相应的excel文件存盘
		FileOutputStream fos = null;
		SXSSFWorkbook workBook = null;
		long t1 = System.currentTimeMillis();
		try {
			File file = new File("E:/hos.xlsx");
			if (file.exists()) {
				file.delete();
			}
			fos = new FileOutputStream(file);

			workBook = new SXSSFWorkbook(100);
			// 在工作薄中创建一工作表
			SXSSFSheet sheet = workBook.createSheet();
			SXSSFRow row = null;
			SXSSFCell code = null;
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 99));
			sheet.addMergedRegion(new CellRangeAddress(1, 99, 0, 0));
			CellStyle cellStyle = workBook.createCellStyle();
			// 设置背景色
			cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT
					.getIndex());
			cellStyle.setFillPattern(FillPatternType.LESS_DOTS);// 设置图案样式

			Font font = workBook.createFont();
			font.setFontName("华文行楷");// 设置字体名称
			font.setFontHeightInPoints((short) 14);// 设置字号
			font.setColor(IndexedColors.RED.getIndex());// 设置字体颜色
			cellStyle.setBorderTop(BorderStyle.THIN);
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setBorderRight(BorderStyle.THIN);
			cellStyle.setBottomBorderColor(IndexedColors.BLUE.getIndex());
			cellStyle.setFont(font);
			for (int i = 0; i < 100; i++) {
				// 在指定的索引处创建一行
				row = sheet.createRow(i);
				for (int j = 0; j < 100; j++) {
					// 在指定的索引处创建一列（单元格）
					// 定义单元格为字符串类型
					code = row.createCell(j, CellType.STRING);
					// 在单元格输入内容
					code.setCellValue(i * j);
				}
			}
			row = sheet.createRow(100);
			for (int j = 1; j < 100; j++) {
				code = row.createCell(j, CellType.STRING);
				code.setCellFormula(String.format("SUM(%s%s:%s%s)", toRadix(j),
						1, toRadix(j), 100));
				System.out.println(String.format("SUM(%s%s:%s%s)", toRadix(j),
						1, toRadix(j), 100));
				code.setCellStyle(cellStyle);
				code.setCellType(CellType.FORMULA);
			}
			sheet.setForceFormulaRecalculation(true);
			workBook.write(fos);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			Toolkit.safeClose(workBook, fos);
			System.out.println("文件生成");
			System.out.println((System.currentTimeMillis() - t1) / 1000);
		}
	}

	public static String toRadix(Integer num) {
		num++;
		String[] array = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
				"K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
				"W", "X", "Y", "Z" };
		int count = 26;
		String out = "";
		if (num / count != 0) {
			out = array[num / count - 1];
			if (num % count == 0) {
				out = out + array[num % count];
			} else {
				out = out + array[num % count - 1];
			}
		} else {
			out = array[num - 1];
		}
		return out;
	}

	interface CallBack {
		public void relase();
	}

	private void prepared() {

	}
}
