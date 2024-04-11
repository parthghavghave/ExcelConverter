package main.java.services;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class service {

	public static void swapRow(Row rowToSwap, Sheet sheet, int lastRowIndex) {

		int rowIndexToSwap = rowToSwap.getRowNum();
		Row lastRow = sheet.getRow(lastRowIndex);

		if (rowToSwap != null && lastRow != null) {

			Row newRow = sheet.createRow(rowIndexToSwap);
			for (int i = 0; i < lastRow.getLastCellNum(); i++) {
				Cell oldCell = lastRow.getCell(i);
				Cell newCell = newRow.createCell(i);
				if (oldCell != null) {
					newCell.setCellValue(oldCell.getStringCellValue());
				}
			}

			Row newLastRow = sheet.createRow(lastRowIndex);
			for (int i = 0; i < rowToSwap.getLastCellNum(); i++) {
				Cell oldCell = rowToSwap.getCell(i);
				Cell newCell = newLastRow.createCell(i);
				if (oldCell != null) {
					newCell.setCellValue(oldCell.getStringCellValue());
				}
			}
		}

	}

	public static Sheet addColumn(int columnIndex, Sheet inputsheet, int outColumnIndex, Sheet outSheet) {

		int rowNum = 0;
		for (Row row : inputsheet) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				Row newRow;
				if (outSheet.getRow(rowNum) == null)
					newRow = outSheet.createRow(rowNum++);
				else
					newRow = outSheet.getRow(rowNum++);
				Cell newCell = newRow.createCell(outColumnIndex);
				newCell.setCellValue(cell.getStringCellValue());
			} else
				break;
		}
		return outSheet;
	}

	public static Sheet addCustomFUPColumn(int columnIndex, Sheet inputsheet, int outColumnIndex, Sheet outSheet) {

		int rowNum = 0;
		for (Row row : inputsheet) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				Row newRow;
				if (outSheet.getRow(rowNum) == null)
					newRow = outSheet.createRow(rowNum++);
				else
					newRow = outSheet.getRow(rowNum++);
				Cell newCell = newRow.createCell(outColumnIndex);
				if (cell.getStringCellValue().contains("/01/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Jan");
				else if (cell.getStringCellValue().contains("/02/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Feb");
				else if (cell.getStringCellValue().contains("/03/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Mar");
				else if (cell.getStringCellValue().contains("/04/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Apr");
				else if (cell.getStringCellValue().contains("/05/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "May");
				else if (cell.getStringCellValue().contains("/06/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Jun");
				else if (cell.getStringCellValue().contains("/07/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Jul");
				else if (cell.getStringCellValue().contains("/08/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Aug");
				else if (cell.getStringCellValue().contains("/09/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Sep");
				else if (cell.getStringCellValue().contains("/10/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Oct");
				else if (cell.getStringCellValue().contains("/11/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Nov");
				else if (cell.getStringCellValue().contains("/12/"))
					newCell.setCellValue(cell.getStringCellValue().substring(0, 2) + "-" + "Dec");
				else
					newCell.setCellValue(cell.getStringCellValue());
			} else
				break;
		}
		return outSheet;
	}

	public static Sheet addCustomNameColumn(int columnIndex, Sheet inputsheet, int outColumnIndex, Sheet outSheet) {

		int rowNum = 0;
		for (Row row : inputsheet) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				Row newRow;
				if (outSheet.getRow(rowNum) == null)
					newRow = outSheet.createRow(rowNum++);
				else
					newRow = outSheet.getRow(rowNum++);
				Cell newCell = newRow.createCell(outColumnIndex);
				if (cell.getStringCellValue().equals("Name"))
					newCell.setCellValue(cell.getStringCellValue());
				else {
					String[] parts = cell.getStringCellValue().split("[.\\s]+");
					if (parts[0].equals("Smt") || parts[0].equals("Sri") || parts[0].equals("Sau")
							|| parts[0].equals("Mrs") || parts[0].equals("Ku") || parts[0].equals("Ms") || parts[0].equals("Shri") || parts[0].length()==1) {
						newCell.setCellValue(parts[3]);
						Cell newCellTwo = newRow.createCell(outColumnIndex + 1);
						newCellTwo.setCellValue(parts[1]);
					} else {
						newCell.setCellValue(parts[2]);
						Cell newCellTwo = newRow.createCell(outColumnIndex + 1);
						newCellTwo.setCellValue(parts[0]);
					}
				}
			} else
				break;
		}
		return outSheet;
	}

	public static Sheet addCustomModeColumn(int columnIndex, Sheet inputsheet, int outColumnIndex, Sheet outSheet) {

		int rowNum = 0;
		for (Row row : inputsheet) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				Row newRow;
				if (outSheet.getRow(rowNum) == null)
					newRow = outSheet.createRow(rowNum++);
				else
					newRow = outSheet.getRow(rowNum++);
				Cell newCell = newRow.createCell(outColumnIndex);
				if (cell.getStringCellValue().equals("Mode"))
					newCell.setCellValue("M");
				else if (cell.getStringCellValue().equals("HLY"))
					newCell.setCellValue("H");
				else if (cell.getStringCellValue().equals("QLY"))
					newCell.setCellValue("Q");
				else if (cell.getStringCellValue().equals("YLY"))
					newCell.setCellValue("Y");
				else
					newCell.setCellValue(cell.getStringCellValue());
			} else
				break;
		}
		return outSheet;
	}

	public static Sheet addCustomPremiumColumn(int columnIndex, Sheet inputsheet, int outColumnIndex, Sheet outSheet) {

		int rowNum = 0;
		for (Row row : inputsheet) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				Row newRow;
				if (outSheet.getRow(rowNum) == null)
					newRow = outSheet.createRow(rowNum++);
				else
					newRow = outSheet.getRow(rowNum++);
				Cell newCell = newRow.createCell(outColumnIndex);
				if (cell.getStringCellValue().equals("Premium(+Tax)"))
					newCell.setCellValue("Premium");
				else
					newCell.setCellValue(
							cell.getStringCellValue().substring(0, cell.getStringCellValue().length() - 2));
			} else
				break;
		}
		return outSheet;
	}

}
