package cn.chenpeng.officetool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OfficeTool {
	public static void main(String[] args) {
		// 获取程序运行路径
		File dir = new File(System.getProperty("user.dir"));
		// 检测程序的文件夹里是否包含这三个文件
		int isExsit = checkFile(dir);
		if (isExsit != 1111) {
			throw new RuntimeException("检测到并非三个文件都存在！");
		}

		// 建立文件对象，根据名字中包含的字段信息匹配哪个文件对应哪个对象！
		File recordFile = null;
		File listFile = null;
		File allListFile = null;
		File[] files = dir.listFiles();

		for (File perfile : files) {
			if (perfile.getName().contains("工作记录") && !perfile.getName().startsWith("~")) {
				recordFile = perfile;
			} else if (perfile.getName().contains("用户") && !perfile.getName().startsWith("~")) {
				listFile = perfile;
			} else if (perfile.getName().contains("员工花名册") && !perfile.getName().startsWith("~")) {
				allListFile = perfile;
			}
		}

		// 将“工作记录名单”文件对象转换成Row集合对象方便后期逐行匹配
		ArrayList<Row> recordAl = convertToList(recordFile);
		// 将“用户”对象转换成Map集合对象，方便后期逐行匹配
		TreeMap<String, String> listMap = convertToMap(listFile);
		// 将上面两者匹配完后的结果放到结果集合中，这个集合中存放的就是在职但是没有ongoing属性的人员ID
		ArrayList<String> resultList = new ArrayList<String>();

		Iterator<String> it = listMap.keySet().iterator();
		while (it.hasNext()) {
			String id = it.next();
			int n = 0;
			for (int i = 0; i < recordAl.size(); i++) {
				Row row = recordAl.get(i);
				if (row.getCell(0).getStringCellValue().equals(id)) {
					if (row.getCell(15).getStringCellValue().equals("ongoing")) {
						n++;
					}
				}
			}

			if (n == 0) {
				resultList.add(id);
			}

		}
		String newFileName = null;
		if (resultList.size() != 0) {

			// 如果结果集的大小不为0，说明在职但是没有ongoing属性的人员ID已经匹配出来，下面就逐一展示出来：
			System.out.println("-----------------------------------------------");
			for (int i = 0; i < resultList.size(); i++) {
				System.out.println("在职无ongoing的人员ID ：" + resultList.get(i));
			}
			System.out.println();

			// 接下来就是把满足条件的人员打印到新的excel表格中！
			Calendar cl = Calendar.getInstance();
			newFileName = "Result" + cl.get(Calendar.YEAR) + "-" + (cl.get(Calendar.MONTH) + 1) + "-"
					+ cl.get(Calendar.DAY_OF_MONTH) + "_" + cl.get(Calendar.HOUR_OF_DAY) + "-" + cl.get(Calendar.MINUTE)
					+ "-" + cl.get(Calendar.SECOND) + ".xlsx";
			File newFile = new File(dir + File.separator + newFileName);
			FileOutputStream fos = null;
			try {
				fos = new FileOutputStream(newFile);
			} catch (FileNotFoundException e1) {
				e1.printStackTrace();
			}
			Workbook newworkbook = new XSSFWorkbook();
			Sheet sheet = newworkbook.createSheet("0");
			// 将“员工花名册”转换成Map集合，方面查询！
			TreeMap<String, Row> allListMap = convertTosMap(allListFile,"在职员工信息",27);
			TreeMap<String, Row> allListMap2 = convertTosMap(allListFile,"离职员工信息",33);
			
			// 将要导出的excel文件第一行标题设置好
			Row row2First = sheet.createRow(0);
			row2First.createCell(0).setCellValue("姓名");
			row2First.createCell(1).setCellValue("工号");
			row2First.createCell(2).setCellValue("部门");
			row2First.createCell(3).setCellValue("岗位");
			row2First.createCell(4).setCellValue("合同公司");
			row2First.createCell(5).setCellValue("工作邮箱");
			row2First.createCell(6).setCellValue("证件号码");
			row2First.createCell(7).setCellValue("手机号");
			row2First.createCell(8).setCellValue("是否在职");

			for (int i = 0; i < resultList.size(); i++) {
				Row row = allListMap.get(resultList.get(i));
				Row row2 = sheet.createRow(i + 1);
				
				if(row==null) {
					row = allListMap2.get(resultList.get(i));
					if(row==null) {
						System.out.println("请注意：[ "+ resultList.get(i) +" ]这个人不在员工花名册内！");
						continue;
					}
					System.out.println("请注意：[ "+ resultList.get(i) +" ]这个人已离职！");
					row2.createCell(0).setCellValue(row.getCell(1).toString());
					row2.createCell(1).setCellValue(row.getCell(2).toString());
					row2.createCell(2).setCellValue(row.getCell(3).toString());
					row2.createCell(3).setCellValue(row.getCell(4).toString());
					row2.createCell(4).setCellValue(row.getCell(5).toString());
					row2.createCell(5).setCellValue(row.getCell(11).toString());
					row2.createCell(6).setCellValue(row.getCell(33).toString());
					row2.createCell(7).setCellValue(row.getCell(50).toString());
					row2.createCell(8).setCellValue("离职");
					continue;
				}
				
				row2.createCell(0).setCellValue(row.getCell(1).toString());
				row2.createCell(1).setCellValue(row.getCell(2).toString());
				row2.createCell(2).setCellValue(row.getCell(3).toString());
				row2.createCell(3).setCellValue(row.getCell(4).toString());
				row2.createCell(4).setCellValue(row.getCell(5).toString());
				row2.createCell(5).setCellValue(row.getCell(11).toString());
				row2.createCell(6).setCellValue(row.getCell(27).toString());
				row2.createCell(7).setCellValue(row.getCell(44).toString());
				row2.createCell(8).setCellValue("在职");
				
			}
			newworkbook.setSheetName(0, "result");
			try {
				newworkbook.write(fos);
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				fos = null;
			}
		}
		System.out.println("-----------------------------------------------");
		System.out.println();
		System.out.println("提取完成，结果已保存至" + newFileName + "中！");
		System.out.println();

	}

	private static TreeMap<String, Row> convertTosMap(File file,String tableName,int columnNum) {
		TreeMap<String, Row> map = new TreeMap<String, Row>();
		try {
			InputStream is = new FileInputStream(file);
			Workbook workBook = WorkbookFactory.create(is);
			is.close();
			Sheet sheet = workBook.getSheet(tableName);
			int rowNum = sheet.getLastRowNum() + 1;

			for (int i = 1; i < rowNum; i++) {
				Row row = sheet.getRow(i);
				String key = row.getCell(columnNum).toString();
				map.put(key, row);
			}
			return map;

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return map;
	}

	private static TreeMap<String, String> convertToMap(File file) {
		TreeMap<String, String> map = new TreeMap<String, String>();

		try {
			InputStream is = new FileInputStream(file);
			Workbook workBook = WorkbookFactory.create(is);
			is.close();
			Sheet sheet = workBook.getSheetAt(0);
			int rowNum = sheet.getLastRowNum() + 1;
			for (int i = 0; i < rowNum; i++) {
				Row row = sheet.getRow(i);
				String key = row.getCell(2).toString();
				String value = row.getCell(16).toString();
				if ("在职".equals(value)) {
					map.put(key, value);
				}
			}
			return map;

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return map;
	}

	private static ArrayList<Row> convertToList(File file) {
		ArrayList<Row> al = new ArrayList<Row>();

		try {
			InputStream is = new FileInputStream(file);
			Workbook workBook = WorkbookFactory.create(is);
			is.close();
			Sheet sheet = workBook.getSheetAt(0);
			int rowNum = sheet.getLastRowNum() + 1;
			for (int i = 0; i < rowNum; i++) {
				Row row = sheet.getRow(i);
				al.add(row);
			}
			return al;

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return al;
	}

	private static int checkFile(File dir) {
		File[] files = dir.listFiles();
		int isRecord = 0;
		int isList = 0;
		int isListAll = 0;

		for (File file1 : files) {
			if (file1.getName().contains("工作记录")) {
				isRecord = 1;
			} else if (file1.getName().contains("用户")) {
				isList = 1;
			} else if (file1.getName().contains("员工花名册")) {
				isListAll = 1;
			}
		}

		return Integer.parseInt("" + 1 + isRecord + isList + isListAll);
	}
}
