import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

/**
 * Created by jdy on 2015/11/29
 */
public class HttpRequest {
	public static String sendGet(String url, String param) {
		String result = "";
		BufferedReader in = null;
		try {
			String urlNameString = url + "?" + param;
			URL realUrl = new URL(urlNameString);
			// �򿪺�URL֮�������
			URLConnection connection = realUrl.openConnection();
			// ����ͨ�õ���������
			connection.setRequestProperty("accept", "*/*");
			connection.setRequestProperty("connection", "Keep-Alive");
			connection.setRequestProperty("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
			// ����ʵ�ʵ�����
			connection.connect();
			// ��ȡ������Ӧͷ�ֶ�
			Map<String, List<String>> map = connection.getHeaderFields();
			// �������е���Ӧͷ�ֶ�
			// for (String key : map.keySet()) {
			// System.out.println(key + "--->" + map.get(key));
			// }
			// ���� BufferedReader����������ȡURL����Ӧ
			in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
			String line;
			while ((line = in.readLine()) != null) {
				result += line;
			}
		} catch (Exception e) {
			// System.out.println("����GET��������쳣��" + e);
			// e.printStackTrace();
		}
		// ʹ��finally�����ر�������
		finally {
			try {
				if (in != null) {
					in.close();
				}
			} catch (Exception e2) {
				// e2.printStackTrace();
			}
		}
		return result;
	}

	public static void main(String[] args) throws Exception {
		String fileToBeRead = "D:/test.xlsx";
		InputStream instream = new FileInputStream(fileToBeRead);
		Workbook wb;
		try {
			wb = new XSSFWorkbook(instream);
		} catch (Exception e) {
			wb = new HSSFWorkbook(instream);
		}
		Sheet sheet = wb.getSheetAt(0);
		for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
			Row row = sheet.getRow(rowNum);
			String url = "http://api.altmetric.com/v1/doi/";
			if (null != row) {
				if (null != row.getCell(3)) {
					String str = row.getCell(3).toString();
					if (StringUtils.isNotBlank(str)) {
						url = url + str;
						String s = HttpRequest.sendGet(url, "");
						String score = "";
						if (StringUtils.isNotBlank(s)) {
                            try {
                                JSONObject jsonObj = new JSONObject(s);
                                score = jsonObj.get("score").toString();
                                System.out.println(jsonObj.get("score"));
                            } catch (Exception e) {
                                System.out.println(url);
                            }
						}
						row.createCell(4).setCellValue(score);
					}
				}
			}
		}
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(fileToBeRead);
			wb.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
