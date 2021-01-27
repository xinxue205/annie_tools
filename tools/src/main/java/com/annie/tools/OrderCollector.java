package com.annie.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import net.sf.json.JSONObject;
//import org.codehaus.jackson.map.ObjectMapper;

public class OrderCollector {
	private static final String FILE_POSTFIX = ".xlsx";
	private static String sourceFileName = "单号"+ FILE_POSTFIX;
	private static String sourceDir = ".";
	static String targetDir = ".";
	static int sourceBeginLine = 2;
	static String[] targetHeader = {"订单号", "收件人姓名", "收件人电话", "快递单号", "单品名称"};
	static Map<String, Integer> kdm_idx = new HashMap<String, Integer>();
	static Map<String, String> kdh_kdm = new HashMap<String, String>();
	static List<String[]>[] allData = new ArrayList[9];
	
	public static void main(String[] args) throws Throwable {
		init();
		readData(); 
		exportData();
	}

	private static void exportData() throws Exception {
		int j = 0;
		for (Entry<String, Integer> e : kdm_idx.entrySet()) {
			List<String[]> data = allData[e.getValue()];
			if(data.size()==0) {
				System.out.println(e.getKey()+" 快递无订单");
				continue;
			}
			XSSFWorkbook workbook = new XSSFWorkbook(); 
			XSSFSheet sheet = workbook.createSheet();
//            XSSFRow row = sheet.createRow(0); 
//            XSSFCell cell = row.createCell(0);
//        	cell.setCellValue("收件人姓名");
//        	cell = row.createCell(1);
//            cell.setCellValue("收件人电话");
//            cell = row.createCell(2);
//            cell.setCellValue("快递单号");
//            cell = row.createCell(3);
//            cell.setCellValue("商品名称");
//            int startRowNo = 1;
			for (int i = 0; i < data.size(); i++) {
				String[] currData = data.get(i);
				XSSFRow row = sheet.createRow(i); 
				XSSFCell cell = row.createCell(0);
	        	cell.setCellValue(currData[0]);
	        	cell = row.createCell(1);
	        	cell.setCellValue(currData[3]);
	        	j++;
			}
			
			//写入文件
            FileOutputStream out = null; 
            try {
            	new File(targetDir+File.separator+"快递").mkdirs();
                out = new FileOutputStream(targetDir+File.separator+"快递"+File.separator+e.getKey()+ FILE_POSTFIX); 
                workbook.write(out); 
            } catch (IOException e1) { 
                e1.printStackTrace(); 
            } finally { 
                try { 
                    out.close(); 
                } catch (IOException e1) { 
                    e1.printStackTrace(); 
                } 
            }
    		System.out.println(e.getKey()+" 快递处理完（共"+ data.size() +"条数据）！");
		}
		System.out.println("此次处理完毕(共"+ j +"条数据)！");
	}
	
	private static void init() throws Exception  {
		System.out.println("注意：1.原数据表格在文件的第一");
		kdm_idx.put("顺丰", 0);
		kdm_idx.put("韵达", 1);
		kdm_idx.put("中通", 2);
		kdm_idx.put("圆通", 3);
		kdm_idx.put("申通", 4);
		kdm_idx.put("京东", 5);
		kdm_idx.put("邮政", 6);
		kdm_idx.put("百世", 7);
		kdm_idx.put("其它", 8);
		
		allData[0] = new ArrayList<String[]>();
		allData[1] = new ArrayList<String[]>();
		allData[2] = new ArrayList<String[]>();
		allData[3] = new ArrayList<String[]>();
		allData[4] = new ArrayList<String[]>();
		allData[5] = new ArrayList<String[]>();
		allData[6] = new ArrayList<String[]>();
		allData[7] = new ArrayList<String[]>();
		allData[8] = new ArrayList<String[]>();
		
		Properties s = getConf();
//			JSONObject jsonObject = JSONObject.fromObject(s);
//			ObjectMapper mapper = new ObjectMapper();
//	        Map readValue = mapper.readValue(s, Map.class);
		sourceFileName = (String) s.getProperty("sourceFileName");
		sourceDir = (String) s.getProperty("sourceDir");
		targetDir = (String) s.getProperty("targetDir");
		
		System.out.println("-----------配置 开始-----------");
		System.out.println(sourceFileName);
		System.out.println(sourceDir);
		System.out.println(targetDir);
		System.out.println("-----------配置 结束-----------");
		
		for (Entry<String, Integer> e : kdm_idx.entrySet()) {
			new File(targetDir+File.separator+"快递"+File.separator+e.getKey()+ FILE_POSTFIX).delete();
			
			String kdh = s.getProperty(e.getKey());
			if (kdh == null) continue;
			if(kdh.indexOf(",")>0) {
				String[] kdhs = kdh.split(",");
				for (int i = 0; i < kdhs.length; i++) {
					kdh_kdm.put(kdhs[i], e.getKey());
				}
			} else {
				kdh_kdm.put(kdh, e.getKey());
			}
		}
	}

	private static void readData() throws Exception {
		FileInputStream fis = new FileInputStream(sourceDir+File.separator+sourceFileName);  
		Workbook wb = WorkbookFactory.create(fis); 
		Sheet sheet = wb.getSheetAt(wb.getNumberOfSheets()-1); 
		System.out.println("当前处理的表格名："+sheet.getSheetName());
		int rowNumbers = sheet.getLastRowNum() + 1;
		List list = new ArrayList();
		for (int row = 0; row < rowNumbers; row++) {
			if(row < sourceBeginLine-1) {
				continue;
			}
			Row r = sheet.getRow(row);
			if(Cell.CELL_TYPE_BLANK == r.getCell(0).getCellType()) {
				continue;
			}
			
			String[] data = new String[6];
			
			for (int col = 0; col < targetHeader.length; col++) {
				Cell cell = r.getCell(col);
				
				String str1 = col==0 ? "" : "\t";
				String val = "";
				if(cell == null) {
					val = "";
				} else {
					
					if( cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						val = NumberToTextConverter.toText(cell.getNumericCellValue());					
					} else {
						val = cell.getStringCellValue();
					}
				}
				str1 += val;
				
				data[col] = val;
			}
			String kdh = data[3];
			String kdhHead = kdh_kdm.get(kdh.substring(0, 2));
			int kdmIndex = 8;
			if(kdhHead != null) {
				kdmIndex = kdm_idx.get(kdhHead);
			} else {
				System.out.println(kdh + " 未知的快递公司，按其它处理");
			}
			List kdxx = allData[kdmIndex];
			kdxx.add(data);
		}
	}

	private static Properties getConf() throws Exception {
		Properties properties = new Properties();
		InputStreamReader inputStream = new InputStreamReader(new FileInputStream("conf.ini"),"UTF-8");
		properties.load(inputStream);
		return properties;
	}
}
