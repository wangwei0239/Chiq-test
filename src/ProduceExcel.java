import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

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
	public static boolean USE_CACHE = false;
	public static String CMPED_INTENT = "";

	public static void main(String[] args) {

		if (args.length < 7) {
			System.out.println("参数为：\"输入文件地址\",\"输出文件地址\",\"访问的url\",\"appid\",\"userid\",\"是否使用cache\",\"需要对比的intent(udf,df,ch)\" ");
			return;
		}

		IN_PATH = args[0];
		OUT_PATH = args[1];
		URL = args[2];
		APP_ID = args[3];
		USER_ID = args[4];
		String useCache = args[5];
		if(useCache.toLowerCase().equals("true")){
			USE_CACHE = true;
		}
		CMPED_INTENT = args[6];

		ExecutorService fixedThreadPool = Executors.newFixedThreadPool(1);
		ExcelUtil readExcel = new ExcelUtil();

		ArrayList<Result> results = new ArrayList<>();
		try {
			ArrayList<Input> questions = (ArrayList<Input>) readExcel.readXlsx(IN_PATH);
			for (Input q : questions) {
				fixedThreadPool.execute(new MyRunnable(results, q));
			}

			fixedThreadPool.shutdown();
			fixedThreadPool.awaitTermination(1, TimeUnit.DAYS);

			System.out.println("------------------------------------------");

			readExcel.writeXlsx(OUT_PATH, results);

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

	public MyRunnable(ArrayList<Result> results, Input input) {
		this.results = results;
		this.input = input;
	}

	@SuppressWarnings("unused")
	@Override
	public void run() {
		Request request = new Request();
		Document document = new Document();
		long startTime = new Date().getTime();
		long timeConsumed = -1;
		document.append("appid", ProduceExcel.APP_ID);// idc
		document.append("cmd", "chat");
		document.append("userid", ProduceExcel.USER_ID);
		document.append("text", input.question);
		if(!ProduceExcel.USE_CACHE){
			document.append("nocache", "1");
		}
		String result = null;

		try {
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
			int userDefaultScore = 0;
			int defaultScore = 0;
			Iterator intentIterator = intents.iterator();

			while (intentIterator.hasNext()) {
				JSONObject intentJson = (JSONObject) intentIterator.next();
				String intentType = intentJson.getString(CATEGORY);
				String value = intentJson.getString("value");
				int score = intentJson.getInt("score");
				String readableIntent = value;

				if (intentType.equals(USER_DEFINE)) {
					if(userDefaultScore < score){
						userDefaultScore = score;
						userDefineIntent = readableIntent;
					}
				} else if (intentType.equals(DEFAULT)) {
					if(defaultScore < score){
						defaultScore = score;
						defaultIntent = readableIntent;
					}
				} else if (intentType.equals(CHANG_HONG)) {
					if(changhongScore < score){
						changhongScore = score;
						changhongIntent = readableIntent;
					}
				}
			}

			Result resultObj = new Result();
			resultObj.expectedDomain = input.expectedDomain;
			resultObj.domain = domain;
			resultObj.expectedDataInsideIntent = input.expectedDataInsideIntent;
			resultObj.intent = intent;
			resultObj.expectedSemantic = input.expectedSemantic;
			resultObj.semantic = semantic;
			resultObj.changhongIntent = changhongIntent;
			resultObj.userDefineIntent = userDefineIntent;
			resultObj.defaultIntent = defaultIntent;
			resultObj.timeConsumed = timeConsumed;
			resultObj.question = input.question;
			resultObj.expectedIntent = input.expectedIntent;
			resultObj.originJson = result;

			synchronized (results) {
				results.add(resultObj);
			}

		} catch (JSONException e) {
			e.printStackTrace();
			Result resultObj = new Result();
			resultObj.question = input.question;
			resultObj.expectedIntent = input.expectedIntent;
			resultObj.expectedDomain = input.expectedIntent;
			resultObj.expectedDataInsideIntent = input.expectedDataInsideIntent;
			resultObj.expectedSemantic = input.expectedSemantic;
			if (result != null) {
				resultObj.originJson = result;
			} else {
				resultObj.originJson = "网络链接错误";
			}
			synchronized (results) {
				results.add(resultObj);
			}
		}

	}

}
