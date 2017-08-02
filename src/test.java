import org.json.JSONException;
import org.json.JSONObject;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;

public class test {
	public static void main(String[] args) {
		String json1 = "[{\"query\":\"$wrongName\",\"spellCheck\":\"$correctName\",\"madeChange\":true,\"stateCode\":0}{\"result\":\"1\"}]";
		String json2 = "[{\"query\":\"$wrongName\",\"spellCheck\":\"$correctName\",\"madeChange\":true,\"stateCode\":0}]";
		System.out.println("Result:"+isValueCorrect(json1, json2));
	}
	
	private static boolean isValueCorrect(String expected, String realValue){
		boolean isJson = false;
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
		
	}
}
