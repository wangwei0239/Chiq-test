import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;
import java.util.Date;
import java.util.Iterator;

import org.bson.Document;

public class Request {
	
//	public static void main(String[] args) throws IOException {
//		Request request = new Request();
//		Document document = new Document();
//		document.append("appid", "6877bc4cf5ad74a6af8d1076e4eec7be");
//		document.append("cmd", "chat");
//		document.append("userid", "0224ACEC80DDB3443D311E09879943DA9");
//		document.append("text", "把声音调高一点");
//		String result = request._sendRequest("http://idc.emotibot.com/api/ApiKey/openapi.php", document);
//		System.out.println(result);
//	}

	public String _sendRequest(String url, Document body) throws IOException {
		// log.info(url);
		Date start = new Date();
		long end = 0;
		URL fullurl = null;
		try {
			fullurl = new URL(url);
		} catch (MalformedURLException e2) {
		}
		PrintWriter out = null;
		URLConnection conn = null;
		conn = (URLConnection) fullurl.openConnection();
		conn.setConnectTimeout(50000);
		conn.setReadTimeout(50000);
		if (body != null) {
			String bod = "";
			Iterator<String> keyit = body.keySet().iterator();
			while (keyit.hasNext()) {
				String key = keyit.next();
				String val = body.getString(key);
				String pair = key + "=" + val;
				bod += pair + "&";
			}
			bod = bod.substring(0, bod.length() - 1);
			System.out.println(bod);
			// String bod=body.toJson();
			bod = new String(bod.toString().getBytes("UTF-8"));
			// conn.setRequestProperty("Content-Length", ""+bod.length());
			conn.setDoOutput(true);
			conn.setDoInput(true);
			out = new PrintWriter(conn.getOutputStream());
			out.print(bod);
			out.flush();
		}
		String line;
		StringBuilder builder = new StringBuilder();
		BufferedReader reader = null;

		InputStream ia = conn.getInputStream();
		reader = new BufferedReader(new InputStreamReader(ia, "UTF-8"));
		while ((line = reader.readLine()) != null) {
			builder.append(line);
		}
		ia.close();
		// try {
		// Thread.sleep(10);
		// } catch (InterruptedException e) {
		// // TODO Auto-generated catch block
		// e.printStackTrace();
		// }
		return builder.toString();
	}

	public String _sendJSONRequest(String url, Document body) throws IOException {
		// log.info(url);
		Date start = new Date();
		long end = 0;
		URL fullurl = null;
		try {
			fullurl = new URL(url);
		} catch (MalformedURLException e2) {
		}
		PrintWriter out = null;
		URLConnection conn = null;
		conn = (URLConnection) fullurl.openConnection();
		conn.setConnectTimeout(1000);
		conn.setReadTimeout(5000);
		// conn.setRequestProperty("Accept", "*/*");
		// conn.setRequestProperty("Accept-Encoding", "gzip/defalte");
		// conn.setRequestProperty("Accept-Language", "zh-CN,zh;q=0.8");
		// conn.setRequestProperty("connection", "Keep-Alive");
		// conn.setRequestProperty("Origin", "Emotibot Regression UI");
		 conn.setRequestProperty("Content-Type", "application/json");
		// conn.setRequestProperty("Content-Type",
		// "application/x-www-form-urlencoded; charset=UTF-8");
		// conn.setRequestProperty("Referer",
		// "http://400.vip.com/WebChat/chat/vchatWeb/webchat_vip.html?isBroswer=1&vendorid=10000&cih_acs_qs_flag=0&cih_acs_qs_content=A0102&platform=pc&cih_source_url=http%3A%2F%2Facs.vip.com%2F");
		// conn.setRequestProperty("Origin", "http://400.vip.com");
		if (body != null) {
			// String bod="";
			// System.out.println(bod);
			String bod = body.toJson();
//			 System.out.println(bod);
			bod = new String(bod.toString().getBytes("UTF-8"));
			// conn.setRequestProperty("Content-Length", ""+bod.length());
			conn.setDoOutput(true);
			conn.setDoInput(true);
			out = new PrintWriter(conn.getOutputStream());
			out.print(bod);
			out.flush();
			out.close();
		}
		String line;
		StringBuilder builder = new StringBuilder();
		BufferedReader reader = null;

		InputStream ia = conn.getInputStream();
		try {
			reader = new BufferedReader(new InputStreamReader(ia));
			while ((line = reader.readLine()) != null) {
				builder.append(line);
			}
			ia.close();
		} catch (Exception e) {
			ia.close();
		}

		

		// try {
		// Thread.sleep(10);
		// } catch (InterruptedException e) {
		// // TODO Auto-generated catch block
		// e.printStackTrace();
		// }
		return builder.toString();
	}
}
