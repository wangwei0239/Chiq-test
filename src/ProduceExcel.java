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

import net.sourceforge.argparse4j.ArgumentParsers;
import net.sourceforge.argparse4j.impl.Arguments;
import net.sourceforge.argparse4j.inf.ArgumentParser;
import net.sourceforge.argparse4j.inf.ArgumentParserException;
import net.sourceforge.argparse4j.inf.ArgumentType;
import net.sourceforge.argparse4j.inf.Namespace;

public class ProduceExcel {

	public static String IN_PATH = "";
	public static String OUT_PATH = "";
	public static String URL = "";
	public static String APP_ID = "";
	public static String USER_ID = "";
	public static boolean USE_CACHE = false;
	public static String CMPED_INTENT = "";
	public static int NUM_OF_THREAD = 1;
	public static boolean IS_CMP_INTENT = true;
	public static boolean IS_CMP_DOMAIN = true;
	public static boolean IS_CMP_SEMANTIC = true;
	public static boolean IS_CMP_DATA_INTENT = true;

	public static void main(String[] args) {
		
		ArgumentParser parser = ArgumentParsers.newArgumentParser("Chiq-test");
		parser.addArgument("--input_file").required(true).help("输入文件地址").metavar("");
		parser.addArgument("--output_file").required(true).help("输出文件地址").metavar("");
		parser.addArgument("--url").required(true).help("测试url").metavar("");
		parser.addArgument("--appid").required(true).help("appid").metavar("");
		parser.addArgument("--user_id").required(true).help("user id").metavar("");
		parser.addArgument("--use_cache").required(true).type(Boolean.class).help("是否使用cache，使用为'true'，不是用为'false'").metavar("");
		parser.addArgument("--cmped_intent").choices("udf","df","ch").setDefault("ch").help("需要对比的intent").metavar("");
		parser.addArgument("--cmp_intent").type(Boolean.class).setDefault(false).help("是否对比intent").metavar("");
		parser.addArgument("--cmp_domain").type(Boolean.class).setDefault(false).help("是否对比domain").metavar("");
		parser.addArgument("--cmp_semantic").type(Boolean.class).setDefault(false).help("是否对比semantic").metavar("");
		parser.addArgument("--cmp_data_intent").type(Boolean.class).setDefault(false).help("是否对比data里面的intent").metavar("");
		parser.addArgument("--thread_num").type(Integer.class).help("线程数").setDefault(1).metavar("");
		
		Namespace namespace = null;
		try {
			namespace = parser.parseArgs(args);
			namespace.toString();
		} catch (ArgumentParserException e1) {
			parser.handleError(e1);
			parser.printHelp();
			System.exit(1);
		}

//		if (args.length < 7) {
//			System.out.println("参数为：\"输入文件地址\",\"输出文件地址\",\"访问的url\",\"appid\",\"userid\",\"是否使用cache\",\"需要对比的intent(udf,df,ch)\" ");
//			return;
//		}
//
//		IN_PATH = args[0];
//		OUT_PATH = args[1];
//		URL = args[2];
//		APP_ID = args[3];
//		USER_ID = args[4];
//		String useCache = args[5];
//		if(useCache.toLowerCase().equals("true")){
//			USE_CACHE = true;
//		}
//		CMPED_INTENT = args[6];
		
		IN_PATH = namespace.getString("input_file");
		OUT_PATH = namespace.getString("output_file");
		URL = namespace.getString("url");
		APP_ID = namespace.getString("appid");
		USER_ID = namespace.getString("user_id");
		USE_CACHE = namespace.getBoolean("use_cache");
		CMPED_INTENT = namespace.getString("cmped_intent");
		NUM_OF_THREAD = namespace.getInt("thread_num");
		IS_CMP_INTENT = namespace.getBoolean("cmp_intent");
		IS_CMP_DOMAIN = namespace.getBoolean("cmp_domain");
		IS_CMP_SEMANTIC = namespace.getBoolean("cmp_semantic");
		IS_CMP_DATA_INTENT = namespace.getBoolean("cmp_data_intent");
		

		ExecutorService fixedThreadPool = Executors.newFixedThreadPool(NUM_OF_THREAD);
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
		if(input == null || input.question == null || input.question.trim().equals("")){
			return;
		}
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
			result = request._sendRequest(ProduceExcel.URL, document);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		try {
			long endTime = new Date().getTime();
			timeConsumed = endTime - startTime;
			System.out.println("Result:"+result);
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
			resultObj.udfScore = userDefaultScore;
			resultObj.dfScore = defaultScore;
			resultObj.chScore = changhongScore;

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
