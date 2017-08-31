import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
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
	
	public List<Input> readXlsx(String path) throws IOException{
		InputStream is = new FileInputStream(path);
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
		List<Input> list = new ArrayList<>();
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		if(xssfSheet == null){
			xssfWorkbook.close();
			return null;
		}
		for(int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++){
			XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			if(xssfRow != null){
				try {
					XSSFCell content = xssfRow.getCell(0);
					XSSFCell expectedIntent = xssfRow.getCell(1);
					XSSFCell expectedDomain = xssfRow.getCell(2);
					XSSFCell expectedDataIntent = xssfRow.getCell(3);
					XSSFCell expectedSemantic = xssfRow.getCell(4);
					Input input = new Input();
					input.question = getValue(content);
					input.expectedIntent = getValue(expectedIntent);
					input.expectedDomain = getValue(expectedDomain);
					input.expectedDataInsideIntent = getValue(expectedDataIntent);
					input.expectedSemantic = getValue(expectedSemantic);
					list.add(input);
				} catch (NullPointerException e) {
					// TODO: handle exception
				}
			}
		}
		xssfWorkbook.close();
		is.close();
		return list;
	}
	
	public void writeXlsx(String path, ArrayList<Result> results) throws IOException{
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
		row0.createCell(1).setCellValue("期待Domian");
		row0.createCell(2).setCellValue("Domian");
		row0.createCell(3).setCellValue("期待Data Intent");
		row0.createCell(4).setCellValue("Data Intent");
		row0.createCell(5).setCellValue("期待Semantic");
		row0.createCell(6).setCellValue("Semantic");
		row0.createCell(7).setCellValue("期待ChanghongIntent");
		row0.createCell(8).setCellValue("Changhong Intent");
		row0.createCell(9).setCellValue("Default");
		row0.createCell(10).setCellValue("User Define");
		row0.createCell(11).setCellValue("Time Consumed");
		row0.createCell(12).setCellValue("Origin Json");
		row0.createCell(13).setCellValue("Result");
		
		int totalNum = results.size();
		int passNum = 0;
		int failNum = 0;
		
		for(Result result : results){
			boolean cmpResult = true;
			XSSFRow row = sheet.createRow(++rowNum);
			row.createCell(0).setCellValue(result.question);
			row.createCell(1).setCellValue(result.expectedDomain);
			XSSFCell domainCell = row.createCell(2);
			domainCell.setCellValue(result.domain);
			if(!isValueCorrect(result.expectedDomain, result.domain)){
				domainCell.setCellStyle(cellStyle);
				cmpResult = false;
			}
			
			row.createCell(3).setCellValue(result.expectedDataInsideIntent);
			XSSFCell intentCell = row.createCell(4);
			intentCell.setCellValue(result.intent);
			if(!isValueCorrect(result.expectedDataInsideIntent, result.intent)){
				intentCell.setCellStyle(cellStyle);
				cmpResult = false;
			}
			
			row.createCell(5).setCellValue(result.expectedSemantic);
			XSSFCell semanticCellrow = row.createCell(6);
			semanticCellrow.setCellValue(result.semantic);
			if(!isValueCorrect(result.expectedSemantic, result.semantic)){
				semanticCellrow.setCellStyle(cellStyle);
				cmpResult = false;
			}
			
			row.createCell(7).setCellValue(result.expectedIntent);
			XSSFCell changhongIntentCell = row.createCell(8);
			changhongIntentCell.setCellValue(result.changhongIntent);
			
			XSSFCell defaultIntentCell = row.createCell(9);
			defaultIntentCell.setCellValue(result.defaultIntent);
			
			XSSFCell userDefineIntent = row.createCell(10);
			userDefineIntent.setCellValue(result.userDefineIntent);
			
			switch (ProduceExcel.CMPED_INTENT) {
			case "udf":
				if(!isValueCorrect(result.expectedIntent, result.userDefineIntent)){
					userDefineIntent.setCellStyle(cellStyle);
					cmpResult = false;
				}
				break;
			
			case "df":
				if(!isValueCorrect(result.expectedIntent, result.defaultIntent)){
					defaultIntentCell.setCellStyle(cellStyle);
					cmpResult = false;
				}
				break;
				
			case "ch":
			default:
				if(!isValueCorrect(result.expectedIntent, result.changhongIntent)){
					changhongIntentCell.setCellStyle(cellStyle);
					cmpResult = false;
				}
				break;
			}
			
			row.createCell(11).setCellValue(result.timeConsumed);
			row.createCell(12).setCellValue(result.originJson);
			XSSFCell cmpCell = row.createCell(13);
			if(cmpResult){
				cmpCell.setCellValue("Pass");
				cmpCell.setCellStyle(correctStyle);
				passNum++;
			}else {
				cmpCell.setCellValue("Fail");
				cmpCell.setCellStyle(cellStyle);
				failNum++;
			}
			
		}
		
		XSSFRow row = sheet.createRow(++rowNum);
		XSSFCell totalCell = row.createCell(0);
		totalCell.setCellValue("Total:"+totalNum);
		
		XSSFCell passCell = row.createCell(1);
		passCell.setCellValue("Pass:"+passNum);
		passCell.setCellStyle(correctStyle);
		
		XSSFCell failCell = row.createCell(2);
		failCell.setCellValue("Fail:"+failNum);
		failCell.setCellStyle(cellStyle);
		
		System.out.println("Test Result:");
		System.out.println("Pass:"+passNum);
		System.out.println("Fail:"+failNum);
		System.out.println("Total:"+totalNum);
		
		OutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}
	
	private boolean isValueCorrect(String expected, String realValue){
		boolean isJson = false;
		try {
			try {
				JSONObject jsonObject = new JSONObject(expected);
				JSONObject jsonObject2 = new JSONObject(realValue);
				isJson = true;
			} catch (JSONException e) {
			}
			
			if(isJson){
				Gson gsonExpected = new GsonBuilder().create();
				JsonElement elementExpected = gsonExpected.toJsonTree(expected);
				Gson gsonRealValue = new GsonBuilder().create();
				JsonElement elementRealValue = gsonExpected.toJsonTree(realValue);
				return elementExpected.equals(elementRealValue);
			}else {
				return expected.equals(realValue);
			}
		} catch (NullPointerException e) {
			return expected == null;
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
