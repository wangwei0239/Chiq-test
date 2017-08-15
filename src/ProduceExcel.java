import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.UUID;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.bson.Document;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class ProduceExcel {

	public static String IN_PATH = "";
	public static String OUT_PATH = "";
	public static String URL = "";
	public static String APP_ID = "";
	public static String USER_ID = "";
//	public static int THREAD_NUM = 1;
	public static boolean USE_CACHE = false;

	public static void main(String[] args) {

		if (args.length < 6) {
			System.out.println("参数为：\"输入文件地址\",\"输出文件地址\",\"访问的url\",\"appid\",\"userid\",\"是否使用缓存\"");
//			System.out.println("参数为：\"输入文件地址\",\"输出文件地址\",\"访问的url\",\"appid\",\"userid\",\"线程数\",\"是否使用缓存\" ");
			return;
		}

		IN_PATH = args[0];
		OUT_PATH = args[1];
		URL = args[2];
		APP_ID = args[3];
		USER_ID = args[4];
		String useCacheString = args[5];
		if(useCacheString.toLowerCase().equals("true")){
			USE_CACHE = true;
		}
//		String numString = args[5];
//		try {
//			THREAD_NUM = Integer.parseInt(numString);
//		} catch (Exception e) {
//			// TODO: handle exception
//		}
		

		ExecutorService fixedThreadPool = Executors.newFixedThreadPool(20);
		ExcelUtil readExcel = new ExcelUtil();

		ArrayList<Result> results = new ArrayList<>();
		try {
			ArrayList<Input> questions = (ArrayList<Input>) readExcel.readXlsx(IN_PATH);
			readExcel.createXlsx(OUT_PATH);
			for (Input q : questions) {
				fixedThreadPool.execute(new MyRunnable(results, q, readExcel));
			}
			
//			int num = 0;
//			
//			long time = System.currentTimeMillis();
//			
//			long duration = 1000 * 2;
//			
//			while((System.currentTimeMillis() - time) < duration){
//				Input q = questions.get(num);
//				num = (++num)%questions.size();
//				Thread.sleep(10);
////				System.out.println("Current Time:"+(System.currentTimeMillis() - time));
//				
//				fixedThreadPool.execute(new MyRunnable(results, q, readExcel));
//			}

			fixedThreadPool.shutdown();
			fixedThreadPool.awaitTermination(1, TimeUnit.DAYS);

			System.out.println("------------------------------------------");

//			readExcel.writeXlsx(OUT_PATH, results);

		} catch (IOException | InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}

class MyRunnable implements Runnable {

	public static final String DATA = "data";
	public static final String TYPE = "type";
	public static final String CUSTOM_DATA = "customdata";
	public static final String DOMAIN = "domain";
	public static final String INTENT = "intent";
	public static final String SEMANTIC = "semantic";
	public static final String USER_DEFINE = "userDefine";
	public static final String DEFAULT = "default";
	public static final String CHANG_HONG = "changhong";
	public static final String CATEGORY = "category";

	private ArrayList<Result> results;
	private Input input;
	private ExcelUtil excelUtil;

	public MyRunnable(ArrayList<Result> results, Input input, ExcelUtil excelUtil) {
		this.results = results;
		this.input = input;
		this.excelUtil = excelUtil;
	}

	@SuppressWarnings("unused")
	@Override
	public void run() {
		Request request = new Request();
		Document document = new Document();
		long startTime = new Date().getTime();
		long timeConsumed = -1;
		// document.append("appid", "5a200ce8e6ec3a6506030e54ac3b970e");//长虹
		// document.append("appid", "6877bc4cf5ad74a6af8d1076e4eec7be");//idc
		document.append("appid", ProduceExcel.APP_ID);// idc
		document.append("cmd", "chat");
		// document.append("userid", "0E2C2D3D11FB8502E8629830869B05CAD");//长虹
		// document.append("userid", "0224ACEC80DDB3443D311E09879943DA9");//idc
//		document.append("userid", ProduceExcel.USER_ID);
		document.append("userid", UUID.randomUUID().toString());
		document.append("text", input.question);
		if(!ProduceExcel.USE_CACHE){
			document.append("nocache", "1");
		}
		
		String result = null;

		try {
			result = request._sendRequest(ProduceExcel.URL, document);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		try {
			// String result =
			// request._sendRequest("http://emotibot.changhong.com/api/ApiKey/openapi.php",
			// document);//长虹

			// System.out.println(result);

			long endTime = new Date().getTime();
			timeConsumed = endTime - startTime;
			JSONObject jsonObject = new JSONObject(result);
			JSONArray datas = jsonObject.getJSONArray(DATA);
			Iterator iterator = datas.iterator();
			JSONObject customData = null;
			while (iterator.hasNext()) {
				JSONObject data = (JSONObject) iterator.next();
				String type = data.getString(TYPE);
				if (type.equals(CUSTOM_DATA)) {
					customData = data;
					break;
				}
			}
			String domain = null;
			String intent = null;
			String semantic =null;
			if (customData != null) {
				JSONObject innerJson = customData.getJSONObject(DATA);
				domain = innerJson.getString(DOMAIN);
				intent = innerJson.getString(INTENT);
				semantic = innerJson.getJSONObject(SEMANTIC).toString();
			}
			JSONArray intents = jsonObject.getJSONArray(INTENT);
			String changhongIntent = "";
			String userDefineIntent = "";
			String defaultIntent = "";
			int changhongScore = 0;

			Iterator intentIterator = intents.iterator();

			while (intentIterator.hasNext()) {
				JSONObject intentJson = (JSONObject) intentIterator.next();
				String intentType = intentJson.getString(CATEGORY);

				// System.out.println("Category:"+intentType);

				String value = intentJson.getString("value");
				int score = intentJson.getInt("score");
				String readableIntent = value;

				if (intentType.equals(USER_DEFINE)) {
					userDefineIntent += readableIntent;
				} else if (intentType.equals(DEFAULT)) {
					defaultIntent += readableIntent;
				} else if (intentType.equals(CHANG_HONG)) {
					if(changhongScore < score){
						changhongScore = score;
						changhongIntent = readableIntent;
					}
				}
			}

			Result resultObj = new Result();
			resultObj.domain = domain;
			resultObj.intent = intent;
			resultObj.semantic = semantic;
			resultObj.changhongIntent = changhongIntent;
			resultObj.userDefineIntent = userDefineIntent;
			resultObj.defaultIntent = defaultIntent;
			resultObj.timeConsumed = timeConsumed;
			resultObj.question = input.question;
			resultObj.expectedIntent = input.expectedIntent;
			resultObj.originJson = result;
			resultObj.changhongScore = changhongScore;

//			synchronized (results) {
//				results.add(resultObj);
//			}
			excelUtil.writeSingleRecordXlsx(ProduceExcel.OUT_PATH, resultObj);

		} catch (JSONException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			Result resultObj = new Result();
			resultObj.question = input.question;
			resultObj.expectedIntent = input.expectedIntent;
			if (result != null) {
				resultObj.originJson = result;
			} else {
				resultObj.originJson = "网络链接错误";
			}
			synchronized (results) {
				results.add(resultObj);
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}

// class Result{
// public String domain;
// public String semantic;
// public long timeConsumed;
// public String intent;
//
// public String changhongIntent;
// public String userDefineIntent;
// public String defaultIntent;
// }
