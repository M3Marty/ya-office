package org.m3m.office.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class MergeTables {

	private static File exists(File file, String s) {
		if (file.exists()) {
			return file;
		}

		System.out.println(s);
		System.exit(1);
		return null;
	}

	public static void main(String[] args) {
		if (args.length != 4
				|| !args[2].equals("on")
				|| !args[3].matches("\\w:\\d+=\\w:\\d+")) {
			System.out.println("Usage: java -jar join.jar <base file> <additional file> on <row>:<position to put>=<row>:<length of row>");
			System.exit(1);
		}

		File base = exists(new File(args[0]), "Base file " + args[0] + " not found");
		File additional = exists(new File(args[1]), "Additional file " + args[1] + " not found");

		String[] columnsAlp = args[3].split("[=:]");
		int baseColumnOn = columnsAlp[0].charAt(0) - 'A';
		int baseColumnTo = Integer.parseInt(columnsAlp[1]);
		int additionColumnOn = columnsAlp[2].charAt(0) - 'A';
		int additionCopyLength = Integer.parseInt(columnsAlp[3]);

		Map<String, String[][]> addition = readAdditionalTable(additional, additionColumnOn, additionCopyLength);
		createSummarySpreadsheet(base, baseColumnOn, baseColumnTo, addition);
	}

	private static void createSummarySpreadsheet(File base, int baseColumnOn,
			int baseColumnTo, Map<String, String[][]> addition) {
		try (FileInputStream fis = new FileInputStream(base);
		     Workbook workbook = WorkbookFactory.create(fis);
			 Workbook summary = new XSSFWorkbook()) {

			Sheet baseSheet = workbook.getSheetAt(0);
			Sheet newSheet = summary.createSheet("Summary");

			Row toInsert = newSheet.createRow(0);
			for (Row baseRow : baseSheet) {
				for (Cell baseCell : baseRow) {
					Cell newCell = toInsert.createCell(baseCell.getColumnIndex());
					newCell.setCellValue(formatter.formatCellValue(baseCell));
				}

				Cell keyCell = baseRow.getCell(baseColumnOn);
				if (keyCell == null) {
					continue;
				}

				String key = formatter.formatCellValue(keyCell);
				if (!addition.containsKey(key)) {
					continue;
				}

				String[][] table = addition.get(key);
				for (int i = 0; i < table.length; i++) {
					for (int j = 0; j < table[i].length; j++) {
						int targetColNum = baseColumnTo + j;
						Cell targetCell = toInsert.createCell(targetColNum);
						targetCell.setCellValue(table[i][j]);
					}
					toInsert = newSheet.createRow(toInsert.getRowNum() + 1);
				}
				toInsert = newSheet.createRow(toInsert.getRowNum() + 1);
			}
			try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
				summary.write(fos);
			}
		} catch(IOException e){
			throw new RuntimeException(e);
		}
	}

		private static DataFormatter formatter = new DataFormatter();

	public static String[] readRow(Row row, int from, int length) {
		String[] data = new String[length];
		if (length == 0) {
			length = row.getLastCellNum();
		}
		int to = from + length;
		for (int i = from; i < to; i++) {
			data[i - from] = Optional.ofNullable(row.getCell(i))
					.map(formatter::formatCellValue).orElseGet(() -> null);
		}
		return data;
	}

	public static Map<String, String[][]> readAdditionalTable(File additional, int column, int copyLength) {
		Map<String, String[][]> data = new HashMap<>();
		try (FileInputStream fis = new FileInputStream(additional);
			Workbook workbook = WorkbookFactory.create(fis)) {

			Sheet sheet = workbook.getSheetAt(0);
			Iterator<Row> it = sheet.iterator();
			it.next();

			String prevKeyValue = null;
			List<String[]> subtable = new ArrayList<>();

			while (it.hasNext()) {
				Row row = it.next();
				Cell key = row.getCell(column);

				Optional<String> keyValue = Optional.ofNullable(key)
						.map(Cell::getStringCellValue);
				if (keyValue.isPresent() && !keyValue.get().isEmpty()) {
					if (!subtable.isEmpty()) {
						data.put(prevKeyValue, subtable.toArray(new String[subtable.size()][]));
						subtable.clear();
					}
					prevKeyValue = keyValue.get();
				}
				subtable.add(readRow(row, column, copyLength));
			}

			return data;
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	public static Map<String, Map<String, Object>> readExcel(String filePath) throws IOException {
		Map<String, Map<String, Object>> data = new HashMap<>();

		try (FileInputStream fis = new FileInputStream(filePath);
		     Workbook workbook = new XSSFWorkbook(fis)) {

			Sheet sheet = workbook.getSheetAt(0); // Читаем первый лист
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				if (row.getRowNum() == 0) continue; // Пропуск заголовка

				Cell keyCell = row.getCell(0); // Ключ из первого столбца
				if (keyCell == null) continue;

				String key = keyCell.getStringCellValue();
				Map<String, Object> record = new HashMap<>();

				for (int i = 1; i < row.getLastCellNum(); i++) {
					Cell cell = row.getCell(i);
					if (cell != null) {
//						record.put("Column" + i, getCellValue(cell));
					}
				}
				data.put(key, record);
			}
		}

		return data;
	}

	public static void writeExcel(String filePath, List<Map<String, Object>> processedData) {
		try (Workbook workbook = new XSSFWorkbook();
		     FileOutputStream fos = new FileOutputStream(filePath)) {

			Sheet sheet = workbook.createSheet("Processed Data");

			// Заголовки
			Row headerRow = sheet.createRow(0);
			headerRow.createCell(0).setCellValue("Key");
			headerRow.createCell(1).setCellValue("Processed Value");

			// Данные
			int rowIndex = 1;
			for (Map<String, Object> record : processedData) {
				Row row = sheet.createRow(rowIndex++);
				row.createCell(0).setCellValue(record.get("Key").toString());
				row.createCell(1).setCellValue(record.get("ProcessedValue").toString());
			}

			workbook.write(fos);
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}
}
