package com.annie.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;

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
	static String[] targetHeaderJD = {"订单号", "收件人姓名", "收件人电话", "快递单号", "单品名称"};
	static Map<String, Integer> kdm_idx = new HashMap<String, Integer>();
	static Map<String, String> kdh_kdm = new HashMap<String, String>();
	static List<String[]>[] allData = new ArrayList[9];
	private static Properties jdConf = new Properties();
	
	public static void main(String[] args) throws Exception {
		String type = "jd";//args[0];
		if("kd".equalsIgnoreCase(type)) {
			System.out.println("-----------快递单处理-----------");
			kdOrder();
		} else if("jd".equalsIgnoreCase(type)) {
			System.out.println("-----------基地单处理-----------");
			jdOrder();
		} else {
			System.out.println("-----------处理代码不存在，程序即将退出-----------");
		}
	}
	
	public static void jdOrder() throws Exception {
		deleteDir(new File(targetDir+File.separator+"基地"));
		initJD();
		HashMap data = readDataJD(); 
		exportDataJD(data);
	}
	
	private static void exportDataJD(HashMap<String, Object> data) {
		for (Entry e : data.entrySet()) {
			Map<String, Object> baseMap = (Map<String, Object>) e.getValue();//基地
			for (Entry e2 : baseMap.entrySet()) {
				List prodList = (List) e2.getValue();
				XSSFWorkbook workbook = new XSSFWorkbook(); 
				XSSFSheet sheet = workbook.createSheet();
				XSSFRow row0 = sheet.createRow(0); 
				row0.createCell(0).setCellValue("序号");
	        	row0.createCell(1).setCellValue("地址");
	        	row0.createCell(2).setCellValue("姓名");
	        	row0.createCell(3).setCellValue("电话");
	        	row0.createCell(4).setCellValue("数量");
	        	row0.createCell(5).setCellValue("品名");
	        	row0.createCell(6).setCellValue("规格");
	        	
				for (int i = 0; i < prodList.size(); i++) {
					String[] strs = (String[]) prodList.get(i);
					System.out.println(e.getKey()+" "+ e2.getKey() +" "+strs[1]+" "+strs[2]);
					XSSFRow row = sheet.createRow(i+1); 
					row.createCell(0).setCellValue(i+1);
		        	row.createCell(1).setCellValue(strs[0]);
		        	row.createCell(2).setCellValue(strs[1]);
		        	row.createCell(3).setCellValue(strs[2]);
		        	row.createCell(4).setCellValue(strs[3]);
		        	row.createCell(5).setCellValue(strs[4]);
		        	row.createCell(6).setCellValue(strs[5]);
				}
				
				//写入文件
		        FileOutputStream out = null; 
		        try {
		        	new File(targetDir+File.separator+"基地").mkdirs();
		            out = new FileOutputStream(targetDir+File.separator+"基地"+File.separator+e2.getKey()+" "+ e.getKey() + FILE_POSTFIX); 
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
			}
		}
//		System.out.println(e.getKey()+" 快递处理完（共"+ data.size() +"条数据）！");
	}

	private static HashMap readDataJD() throws Exception {
		FileInputStream fis = new FileInputStream(sourceDir+File.separator+"订单"+ FILE_POSTFIX);  
		Workbook wb = WorkbookFactory.create(fis); 
		Sheet sheet = wb.getSheetAt(wb.getNumberOfSheets()-1); 
		System.out.println("当前处理的表格名："+sheet.getSheetName());
		int rowNumbers = sheet.getLastRowNum() + 1;
		HashMap allMap = new HashMap();
		for (int row = 0; row < rowNumbers; row++) {
			if(row < sourceBeginLine-1) {
				continue;
			}
			
			Row r = sheet.getRow(row);
			if(Cell.CELL_TYPE_BLANK == r.getCell(0).getCellType()) {
				continue;
			}
			
			String addr = getValidCellContent(r.getCell(1));
			String name = getValidCellContent(r.getCell(2));
			String phone = getValidCellContent(r.getCell(3));
			String count = getValidCellContent(r.getCell(4));
			String product = getValidCellContent(r.getCell(5));
			String spec = getValidCellContent(r.getCell(6));
			
			boolean findOrNot = false;
			for (Entry e : jdConf.entrySet()) {
				String findName = (String) e.getKey();//产品名
				if(product.indexOf(findName)!=-1) {//找到基地
					String baseName = (String) e.getValue();//基地名
					Map baseMap = (Map) allMap.get(baseName);
					String[] prodInfo = {addr, name, phone, count, product, spec};
					if(baseMap==null) { //新加的基地
						baseMap = new HashMap();
						allMap.put(baseName, baseMap);//基地下加产品
						List prodList = new ArrayList();
						baseMap.put(findName, prodList);//产品下加订单
					} else {//已有的基地
						List prodList = (List) baseMap.get(findName);
						if(prodList==null) {//新加的产品
							prodList = new ArrayList();
							baseMap.put(findName, prodList);//产品下加订单
						} 
					}
					
					List prodList = (List) baseMap.get(findName);
					prodList.add(prodInfo);
					findOrNot = true;
				}
			}
			
			if(!findOrNot) {
				System.out.println("第" + (row+1) + "行的产品（"+ product +"）未找到对应基地");
			}
		}
		return allMap;
	}

	private static String getValidCellContent(Cell cell) {
		String val = "";
		if(cell != null) {
			if( cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				val = NumberToTextConverter.toText(cell.getNumericCellValue());					
			} else {
				val = cell.getStringCellValue();
			}
		}
		return val;
	}

	public static void kdOrder() throws Exception {
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
	
	private static void initJD() throws Exception  {
		InputStreamReader inputStream = new InputStreamReader(new FileInputStream("jd.ini"),"UTF-8");
		jdConf.load(inputStream);
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
		sourceFileName = (String) s.getProperty("sourceFileName", "单号"+FILE_POSTFIX);
		sourceDir = (String) s.getProperty("sourceDir", ".");
		targetDir = (String) s.getProperty("targetDir", ".");
		
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
	
	private static boolean deleteDir(File dir) {
        if (dir.isDirectory()) {
            String[] children = dir.list();
            for (int i=0; i<children.length; i++) {
                boolean success = deleteDir(new File(dir, children[i]));
                if (!success) {
                    return false;
                }
            }
        }
        return dir.delete();
    }
}
