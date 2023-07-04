package com.springboot.excel.Helper;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import com.springboot.excel.Entity.Excel;

public class ExcelHelper {

	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	static String[] HEADERs = { "Id", "Title", "Description", "Published" };
	static String SHEET = "Data";

	public static boolean hasExcelFormat(MultipartFile file) {
		if (!TYPE.equals(file.getContentType())) {
			return false;
		}
		return true;
	}

	public static List<Excel> excelToXCL(InputStream is) {
		try {
			Workbook workbook = new XSSFWorkbook(is);
			Sheet sheet = workbook.getSheet(SHEET);
			Iterator<Row> rows = sheet.iterator();

			List<Excel> excelList = new ArrayList<>();

			int rowNumber = 0;
			while (rows.hasNext()) {
				Row currentRow = rows.next();

				// skip header
				if (rowNumber == 0) {
					rowNumber++;
					continue;
				}

				Iterator<Cell> cellsInRow = currentRow.iterator();

				Excel excel = new Excel();

				int cellIdx = 0;
				while (cellsInRow.hasNext()) {
					Cell currentCell = cellsInRow.next();

					switch (cellIdx) {
					case 0:
						excel.setId((long) currentCell.getNumericCellValue());
						break;

					case 1:
						excel.setTitle(currentCell.getStringCellValue());
						break;

					case 2:
						excel.setDescription(currentCell.getStringCellValue());
						break;

					case 3:
						excel.setPublished(currentCell.getBooleanCellValue());
						break;

					default:
						break;
					}

					cellIdx++;
				}

				excelList.add(excel);
			}

			workbook.close();

			return excelList;
		} catch (IOException e) {
			throw new RuntimeException("Fail to parse Excel file: " + e.getMessage());
		}
	}
}
