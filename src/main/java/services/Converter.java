package main.java.services;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import main.resource.constants;

public class Converter {

	public static void performConversion(String inputFile) throws EncryptedDocumentException, IOException {
		FileInputStream fis = new FileInputStream(inputFile);
		Workbook workbook = WorkbookFactory.create(fis);
		Sheet sheet = workbook.getSheetAt(0);

		Workbook workbookOut = new HSSFWorkbook();
		Sheet outSheet = workbookOut.createSheet("Final Sheet");
		Sheet tempSheet = workbookOut.createSheet("Sheet with all mode");

		for (Cell cell : sheet.getRow(0)) {
			if (cell.getStringCellValue().equals("SN")) {
				service.addColumn(cell.getColumnIndex(), sheet, constants.SN_COLUMN_INDEX, tempSheet);
			} else if (cell.getStringCellValue().equals("F.U.P.")) {
				service.addCustomFUPColumn(cell.getColumnIndex(), sheet, constants.FUP_COLUMN_INDEX, tempSheet);
			} else if (cell.getStringCellValue().equals("D.O.C.")) {
				service.addColumn(cell.getColumnIndex(), sheet, constants.DOC_COLUMN_INDEX, tempSheet);
			} else if (cell.getStringCellValue().equals("Mode")) {
				service.addCustomModeColumn(cell.getColumnIndex(), sheet, constants.MODE_COLUMN_INDEX, tempSheet);
			} else if (cell.getStringCellValue().equals("Policy No.")) {
				service.addColumn(cell.getColumnIndex(), sheet, constants.POLICYNO_COLUMN_INDEX, tempSheet);
			} else if (cell.getStringCellValue().equals("Premium(+Tax)")) {
				service.addCustomPremiumColumn(cell.getColumnIndex(), sheet, constants.PREMIUM_COLUMN_INDEX, tempSheet);
			} else if (cell.getStringCellValue().equals("Name")) {
				service.addCustomNameColumn(cell.getColumnIndex(), sheet, constants.NAME_COLUMN_INDEX, tempSheet);
			}
		}

		int newRowIndex = 0;
		for (int i = tempSheet.getFirstRowNum(); i <= tempSheet.getLastRowNum(); i++) {
			Row row = tempSheet.getRow(i);
			if (!row.getCell(constants.MODE_COLUMN_INDEX).getStringCellValue().equals("Mly")) {
				Row newRow = outSheet.createRow(newRowIndex++);
				for (Cell cell : row) {
					Cell newCell = newRow.createCell(cell.getColumnIndex());
					newCell.setCellValue(cell.getStringCellValue());
				}
			}
		}

		int startRow = 1;
		int endRow = outSheet.getLastRowNum();
		List<Row> rows = new ArrayList<>();
		for (int i = startRow; i <= endRow; i++) {
			rows.add(outSheet.getRow(i));
		}
		Collections.sort(rows, new Comparator<Row>() {
			@Override
			public int compare(Row row1, Row row2) {
				Cell cell1 = row1.getCell(constants.NAME_COLUMN_INDEX);
				Cell cell2 = row2.getCell(constants.NAME_COLUMN_INDEX);
				String value1 = (cell1 == null) ? "" : cell1.getStringCellValue();
				String value2 = (cell2 == null) ? "" : cell2.getStringCellValue();
				return value1.compareTo(value2);
			}
		});
		int rowNum = startRow;
		for (Row sortedRow : rows) {
			Row newRow = outSheet.createRow(rowNum++);
			for (int i = 0; i < sortedRow.getLastCellNum(); i++) {
				Cell cell = newRow.createCell(i);
				Cell originalCell = sortedRow.getCell(i);
				if (originalCell != null) {
					cell.setCellValue(originalCell.getStringCellValue());
				}
			}
		}

		for (int i = 1; i <= outSheet.getLastRowNum(); i++) {
			Row row = outSheet.getRow(i);
			if (row != null) {
				Cell cell = row.createCell(constants.SN_COLUMN_INDEX);
				cell.setCellValue(i);
			}
		}

		// Styling
		outSheet.setColumnWidth(constants.SN_COLUMN_INDEX, constants.SN_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.FUP_COLUMN_INDEX, constants.FUP_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.DOC_COLUMN_INDEX, constants.DOC_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.MODE_COLUMN_INDEX, constants.MODE_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.POLICYNO_COLUMN_INDEX, constants.POLICYNO_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.PREMIUM_COLUMN_INDEX, constants.PREMIUM_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.NAME_COLUMN_INDEX, constants.NAME_COLUMN_WIDTH);
		outSheet.setColumnWidth(constants.NAMETWO_COLUMN_INDEX, constants.NAMETWO_COLUMN_WIDTH);

		CellStyle rowStyle = workbookOut.createCellStyle();
		Font font = workbookOut.createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 12);
		rowStyle.setFont(font);
		rowStyle.setBorderTop(BorderStyle.THIN);
		rowStyle.setBorderBottom(BorderStyle.THIN);
		rowStyle.setBorderLeft(BorderStyle.THIN);
		rowStyle.setBorderRight(BorderStyle.THIN);

		CellStyle noRightBorderStyle = workbookOut.createCellStyle();
		noRightBorderStyle.cloneStyleFrom(rowStyle);
		noRightBorderStyle.setBorderRight(BorderStyle.NONE);

		CellStyle noLeftBorderStyle = workbookOut.createCellStyle();
		noLeftBorderStyle.cloneStyleFrom(rowStyle);
		noLeftBorderStyle.setBorderLeft(BorderStyle.NONE);

		CellStyle rightAlignedCellStyle = workbookOut.createCellStyle();
		rightAlignedCellStyle.cloneStyleFrom(rowStyle);
		rightAlignedCellStyle.setAlignment(HorizontalAlignment.RIGHT);

		for (int i = 1; i <= outSheet.getLastRowNum(); i++) {
			Row row = outSheet.getRow(i);
			for (int j = constants.SN_COLUMN_INDEX; j <= constants.NAMETWO_COLUMN_INDEX; j++) {
				Cell cell = row.getCell(j);
				if (j == constants.NAME_COLUMN_INDEX) {
					cell.setCellStyle(noRightBorderStyle);
				} else if (j == constants.PREMIUM_COLUMN_INDEX) {
					cell.setCellStyle(rightAlignedCellStyle);
				} else if (j == constants.NAMETWO_COLUMN_INDEX) {
					cell.setCellStyle(noLeftBorderStyle);
				} else {
					cell.setCellStyle(rowStyle);
				}
			}
		}

		CellStyle extraStyle = workbookOut.createCellStyle();
		Font font2 = workbookOut.createFont();
		font2.setBold(true);
		font2.setColor(IndexedColors.WHITE.getIndex());
		font2.setFontName("Arial");
		font2.setFontHeightInPoints((short) 12);
		extraStyle.setFont(font2);
		extraStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		extraStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		extraStyle.setAlignment(HorizontalAlignment.CENTER);
		extraStyle.setBorderTop(BorderStyle.THIN);
		extraStyle.setBorderBottom(BorderStyle.THIN);
		extraStyle.setBorderLeft(BorderStyle.THIN);
		extraStyle.setBorderRight(BorderStyle.THIN);
		Row firstRow = outSheet.getRow(0);
		if (firstRow != null) {
			for (Cell cell : firstRow) {
				cell.setCellStyle(extraStyle);
			}
		}

		outSheet.addMergedRegion(
				new CellRangeAddress(0, 0, constants.NAME_COLUMN_INDEX, constants.NAMETWO_COLUMN_INDEX));

		FileOutputStream fos = new FileOutputStream(constants.OUTPPUT_FILE_LOCATION);
		workbookOut.write(fos);

		workbookOut.close();
		workbook.close();
		fis.close();
		fos.close();
	}
}
