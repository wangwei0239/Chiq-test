import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;

public class ExcelUtil {

	private static int rowNum = 1;

	// public static void main(String[] args) {
	// ExcelUtil readExcel = new ExcelUtil();
	// try {
	// List<String> list =
	// readExcel.readXlsx("/Users/wangwei/Documents/changhong/testset.xlsx");
	// System.out.println(list);
	// } catch (IOException e) {
	// // TODO Auto-generated catch block
	// e.printStackTrace();
	// }
	// }

	public List<Input> readXlsx(String path) throws IOException {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
		List<Input> list = new ArrayList<>();
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		if (xssfSheet == null) {
			xssfWorkbook.close();
			return null;
		}
		for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
			XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			if (xssfRow != null) {
				XSSFCell content = xssfRow.getCell(0);
				XSSFCell expectedIntent = xssfRow.getCell(1);
				XSSFCell expectedDomain = xssfRow.getCell(2);
				XSSFCell expectedDataIntent = xssfRow.getCell(3);
				XSSFCell expectedSemantic = xssfRow.getCell(4);
				Input input = new Input();
				input.question = getValue(content);
				input.expectedIntent = getValue(expectedIntent);
				list.add(input);
			}
		}
		xssfWorkbook.close();
		is.close();
		return list;
	}

	public void createXlsx(String path) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet();
		XSSFRow row0 = sheet.createRow(0);
		row0.createCell(0).setCellValue("Question");
		row0.createCell(1).setCellValue("Domian");
		row0.createCell(2).setCellValue("Data Intent");
		row0.createCell(3).setCellValue("Semantic");
		row0.createCell(4).setCellValue("期待ChanghongIntent");
		row0.createCell(5).setCellValue("Changhong Intent");
		row0.createCell(6).setCellValue("Changhong Score");
		row0.createCell(7).setCellValue("Default");
		row0.createCell(8).setCellValue("User Define");
		row0.createCell(9).setCellValue("Time Consumed");
		row0.createCell(10).setCellValue("Origin Json");
		row0.createCell(11).setCellValue("Result");

		OutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}

	public synchronized void writeSingleRecordXlsx(String path, Result result) throws IOException, EncryptedDocumentException, InvalidFormatException {
		 InputStream is = new FileInputStream(path);
		 XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(is);
		
		System.out.println("RowNum:"+rowNum);

		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFillForegroundColor(new XSSFColor(Color.RED));

		XSSFCellStyle correctStyle = workbook.createCellStyle();
		correctStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		correctStyle.setFillForegroundColor(new XSSFColor(Color.GREEN));
		XSSFSheet sheet = workbook.getSheetAt(0);

		boolean cmpResult = true;
		XSSFRow row = sheet.createRow(rowNum++);
		row.createCell(0).setCellValue(result.question);
		XSSFCell domainCell = row.createCell(1);
		domainCell.setCellValue(result.domain);
		// if(!isValueCorrect(result.expectedDomain, result.domain)){
		// domainCell.setCellStyle(cellStyle);
		// cmpResult = false;
		// }

		XSSFCell intentCell = row.createCell(2);
		intentCell.setCellValue(result.intent);
		// if(!isValueCorrect(result.expectedDataInsideIntent,
		// result.intent)){
		// intentCell.setCellStyle(cellStyle);
		// cmpResult = false;
		// }

		XSSFCell semanticCellrow = row.createCell(3);
		semanticCellrow.setCellValue(result.semantic);
		// if(!isValueCorrect(result.expectedSemantic, result.semantic)){
		// semanticCellrow.setCellStyle(cellStyle);
		// cmpResult = false;
		// }

		row.createCell(4).setCellValue(result.expectedIntent);
		XSSFCell changhongIntentCell = row.createCell(5);
		changhongIntentCell.setCellValue(result.changhongIntent);
		if (!isValueCorrect(result.expectedIntent, result.changhongIntent)) {
			changhongIntentCell.setCellStyle(cellStyle);
			cmpResult = false;
		}

		row.createCell(6).setCellValue(result.changhongScore);
		row.createCell(7).setCellValue(result.defaultIntent);
		row.createCell(8).setCellValue(result.userDefineIntent);
		row.createCell(9).setCellValue(result.timeConsumed);
		row.createCell(10).setCellValue(result.originJson);
		XSSFCell cmpCell = row.createCell(11);
		if (cmpResult) {
			cmpCell.setCellValue("Pass");
			cmpCell.setCellStyle(correctStyle);
		} else {
			cmpCell.setCellValue("Fail");
			cmpCell.setCellStyle(cellStyle);
		}

		OutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}

	public void writeXlsx(String path, ArrayList<Result> results) throws IOException {
		// InputStream is = new FileInputStream(path);
		// XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
		// XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		// if(xssfSheet == null){
		// xssfSheet
		// }
		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFillForegroundColor(new XSSFColor(Color.RED));

		XSSFCellStyle correctStyle = workbook.createCellStyle();
		correctStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		correctStyle.setFillForegroundColor(new XSSFColor(Color.GREEN));

		int rowNum = 0;
		XSSFSheet sheet = workbook.createSheet();
		XSSFRow row0 = sheet.createRow(0);
		row0.createCell(0).setCellValue("Question");
		row0.createCell(1).setCellValue("Domian");
		row0.createCell(2).setCellValue("Data Intent");
		row0.createCell(3).setCellValue("Semantic");
		row0.createCell(4).setCellValue("期待ChanghongIntent");
		row0.createCell(5).setCellValue("Changhong Intent");
		row0.createCell(6).setCellValue("Changhong Score");
		row0.createCell(7).setCellValue("Default");
		row0.createCell(8).setCellValue("User Define");
		row0.createCell(9).setCellValue("Time Consumed");
		row0.createCell(10).setCellValue("Origin Json");
		row0.createCell(11).setCellValue("Result");

		int totalNum = results.size();
		int passNum = 0;
		int failNum = 0;

		for (Result result : results) {
			boolean cmpResult = true;
			XSSFRow row = sheet.createRow(++rowNum);
			row.createCell(0).setCellValue(result.question);
			XSSFCell domainCell = row.createCell(1);
			domainCell.setCellValue(result.domain);
			// if(!isValueCorrect(result.expectedDomain, result.domain)){
			// domainCell.setCellStyle(cellStyle);
			// cmpResult = false;
			// }

			XSSFCell intentCell = row.createCell(2);
			intentCell.setCellValue(result.intent);
			// if(!isValueCorrect(result.expectedDataInsideIntent,
			// result.intent)){
			// intentCell.setCellStyle(cellStyle);
			// cmpResult = false;
			// }

			XSSFCell semanticCellrow = row.createCell(3);
			semanticCellrow.setCellValue(result.semantic);
			// if(!isValueCorrect(result.expectedSemantic, result.semantic)){
			// semanticCellrow.setCellStyle(cellStyle);
			// cmpResult = false;
			// }

			row.createCell(4).setCellValue(result.expectedIntent);
			XSSFCell changhongIntentCell = row.createCell(5);
			changhongIntentCell.setCellValue(result.changhongIntent);
			if (!isValueCorrect(result.expectedIntent, result.changhongIntent)) {
				changhongIntentCell.setCellStyle(cellStyle);
				cmpResult = false;
			}

			row.createCell(6).setCellValue(result.changhongScore);
			row.createCell(7).setCellValue(result.defaultIntent);
			row.createCell(8).setCellValue(result.userDefineIntent);
			row.createCell(9).setCellValue(result.timeConsumed);
			row.createCell(10).setCellValue(result.originJson);
			XSSFCell cmpCell = row.createCell(11);
			if (cmpResult) {
				cmpCell.setCellValue("Pass");
				cmpCell.setCellStyle(correctStyle);
				passNum++;
			} else {
				cmpCell.setCellValue("Fail");
				cmpCell.setCellStyle(cellStyle);
				failNum++;
			}

		}

		XSSFRow row = sheet.createRow(++rowNum);
		XSSFCell totalCell = row.createCell(0);
		totalCell.setCellValue("Total:" + totalNum);

		XSSFCell passCell = row.createCell(1);
		passCell.setCellValue("Pass:" + passNum);
		passCell.setCellStyle(correctStyle);

		XSSFCell failCell = row.createCell(2);
		failCell.setCellValue("Fail:" + failNum);
		failCell.setCellStyle(cellStyle);

		OutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}

	private boolean isValueCorrect(String expected, String realValue) {
		boolean isJson = false;
		try {
			JSONObject jsonObject = new JSONObject(expected);
			JSONObject jsonObject2 = new JSONObject(realValue);
			isJson = true;
		} catch (JSONException e) {
		}

		if (isJson) {
			Gson gsonExpected = new GsonBuilder().create();
			JsonElement elementExpected = gsonExpected.toJsonTree(expected);
			Gson gsonRealValue = new GsonBuilder().create();
			JsonElement elementRealValue = gsonExpected.toJsonTree(realValue);
			return elementExpected.equals(elementRealValue);
		} else {
			return expected.equals(realValue);
		}

	}

	@SuppressWarnings({ "deprecation", "static-access" })
	private String getValue(XSSFCell xssfRow) {
		if (xssfRow.getCellType() == xssfRow.CELL_TYPE_BOOLEAN) {
			return String.valueOf(xssfRow.getBooleanCellValue());
		} else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_NUMERIC) {
			return String.valueOf(xssfRow.getNumericCellValue());
		} else {
			return String.valueOf(xssfRow.getStringCellValue());
		}
	}
}
