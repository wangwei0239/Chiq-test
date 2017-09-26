import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;

import com.google.gson.JsonObject;
import com.google.gson.JsonParser;


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
					int columnId = 0;
					XSSFCell content = xssfRow.getCell(columnId++);
					XSSFCell expectedIntent = null;
					XSSFCell expectedDomain = null;
					XSSFCell expectedDataIntent = null;
					XSSFCell expectedSemantic = null;
					
					if(ProduceExcel.IS_CMP_DOMAIN){
						expectedDomain = xssfRow.getCell(columnId++);
					}
					
					if(ProduceExcel.IS_CMP_DATA_INTENT){
						expectedDataIntent = xssfRow.getCell(columnId++);
					}
					if(ProduceExcel.IS_CMP_SEMANTIC){
						expectedSemantic = xssfRow.getCell(columnId++);
					}
					
					if(ProduceExcel.IS_CMP_INTENT){
						expectedIntent = xssfRow.getCell(columnId++);
					}
					
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
		
		int columnId = 0;
		
		row0.createCell(columnId++).setCellValue("Question");
		if(ProduceExcel.IS_CMP_DOMAIN){
			row0.createCell(columnId++).setCellValue("期待Domian");
		}
		row0.createCell(columnId++).setCellValue("Domian");
		if(ProduceExcel.IS_CMP_DATA_INTENT){
			row0.createCell(columnId++).setCellValue("期待Data Intent");
		}
		row0.createCell(columnId++).setCellValue("Data Intent");
		if(ProduceExcel.IS_CMP_SEMANTIC){
			row0.createCell(columnId++).setCellValue("期待Semantic");
		}
		row0.createCell(columnId++).setCellValue("Semantic");
		if(ProduceExcel.IS_CMP_INTENT){
			row0.createCell(columnId++).setCellValue("期待Intent");
		}
		row0.createCell(columnId++).setCellValue("Intent");
		row0.createCell(columnId++).setCellValue("Changhong Intent Score");
		row0.createCell(columnId++).setCellValue("Default");
		row0.createCell(columnId++).setCellValue("Default Intent Score");
		row0.createCell(columnId++).setCellValue("User Define");
		row0.createCell(columnId++).setCellValue("User Define Intent Score");
		row0.createCell(columnId++).setCellValue("Time Consumed");
		row0.createCell(columnId++).setCellValue("Origin Json");
		row0.createCell(columnId++).setCellValue("Result");
		
		int totalNum = results.size();
		int passNum = 0;
		int failNum = 0;
		
		
		
		for(Result result : results){
			columnId = 0;
			boolean cmpResult = true;
			XSSFRow row = sheet.createRow(++rowNum);
			row.createCell(columnId++).setCellValue(result.question);
			if(ProduceExcel.IS_CMP_DOMAIN){
				row.createCell(columnId++).setCellValue(result.expectedDomain);
			}
			XSSFCell domainCell = row.createCell(columnId++);
			domainCell.setCellValue(result.domain);
			if(!isValueCorrect(result.expectedDomain, result.domain)&&ProduceExcel.IS_CMP_DOMAIN){
				domainCell.setCellStyle(cellStyle);
				cmpResult = false;
			}
			if(ProduceExcel.IS_CMP_DATA_INTENT){
				row.createCell(columnId++).setCellValue(result.expectedDataInsideIntent);
			}
			XSSFCell intentCell = row.createCell(columnId++);
			intentCell.setCellValue(result.intent);
			if(!isValueCorrect(result.expectedDataInsideIntent, result.intent)&&ProduceExcel.IS_CMP_DATA_INTENT){
				intentCell.setCellStyle(cellStyle);
				cmpResult = false;
			}
			if(ProduceExcel.IS_CMP_SEMANTIC){
				row.createCell(columnId++).setCellValue(result.expectedSemantic);
			}
			XSSFCell semanticCellrow = row.createCell(columnId++);
			semanticCellrow.setCellValue(result.semantic);
			if(!isValueCorrect(result.expectedSemantic, result.semantic)&&ProduceExcel.IS_CMP_SEMANTIC){
				semanticCellrow.setCellStyle(cellStyle);
				cmpResult = false;
			}
			if(ProduceExcel.IS_CMP_INTENT){
				row.createCell(columnId++).setCellValue(result.expectedIntent);
			}
			XSSFCell changhongIntentCell = row.createCell(columnId++);
			changhongIntentCell.setCellValue(result.changhongIntent);
			
			XSSFCell changhongIntentScoreCell = row.createCell(columnId++);
			changhongIntentScoreCell.setCellValue(result.chScore);
			
			XSSFCell defaultIntentCell = row.createCell(columnId++);
			defaultIntentCell.setCellValue(result.defaultIntent);
			
			XSSFCell defaultIntentScoreCell = row.createCell(columnId++);
			defaultIntentScoreCell.setCellValue(result.dfScore);
			
			XSSFCell userDefineIntent = row.createCell(columnId++);
			userDefineIntent.setCellValue(result.userDefineIntent);
			
			XSSFCell udfIntentScoreCell = row.createCell(columnId++);
			udfIntentScoreCell.setCellValue(result.udfScore);
			
			if(ProduceExcel.IS_CMP_INTENT){
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
			}
			
			row.createCell(columnId++).setCellValue(result.timeConsumed);
			row.createCell(columnId++).setCellValue(result.originJson);
			XSSFCell cmpCell = row.createCell(columnId++);
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
				JsonParser parser = new JsonParser();
                JsonObject elementRealValue = (JsonObject) parser.parse(realValue);
                JsonObject elementExpected = (JsonObject) parser.parse(expected);
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
		if (xssfRow == null){
			return null;
		} else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_BOOLEAN) {
			return String.valueOf(xssfRow.getBooleanCellValue());
		} else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_NUMERIC) {
			return String.valueOf(xssfRow.getNumericCellValue());
		} else {
			return String.valueOf(xssfRow.getStringCellValue());
		}
	}
}
