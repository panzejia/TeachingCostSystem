package cn.iflin.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FCE_one {
	public static void main(String[] args) throws IOException   {
		String filePath = "F:\\Project\\16项目组\\管院近五年课程\\2015-2016学年\\2015-2016-1必修课.xlsx";
		FileInputStream excelFile = new FileInputStream(
				new File(filePath));
		XSSFWorkbook  workbook = new XSSFWorkbook(excelFile);
		GetCellNum(workbook);
	}
	
	/*
	 * 获取excel文件内所需数据列数
	 */
	private static void GetCellNum(XSSFWorkbook  workbook) throws IOException{
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFRow row = sheet.getRow(2);
		int cellNum,cellAllNum = row.getLastCellNum();
		String cellValue=null;
		for(cellNum=0;cellNum<cellAllNum;cellNum++){
			cellValue = row.getCell(cellNum).getStringCellValue();	//获取单元格内的值
			switch (cellValue){	//对取到的值进行匹配 将符合的列数传给OutExcel函数
			case "已选": 
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "课程名称":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "教师姓名":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "课程性质":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "学分":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			default:
				break;
			}
		}
	}
	/*
	 * 获取excel文件内所需数据列数中的所有数据
	 * 需修改
	 */
	@SuppressWarnings("resource")
	private static  void OutExcel(int cellNum,String cellFirstValue) throws IOException{
		/*获取旧Excel文件*/
		String filePath = "F:\\Project\\16项目组\\管院近五年课程\\2015-2016学年\\2015-2016-1必修课.xlsx";
		FileInputStream excelFile = new FileInputStream(
				new File(filePath));
		XSSFWorkbook  workbook = new XSSFWorkbook(excelFile);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		XSSFRow row ;
		String cellStrValue = null;
		double cellNumValue;
		int rowNum,rowAllNum=sheet.getLastRowNum(); //取行数值
		for(rowNum=3;rowNum<rowAllNum;rowNum++){
			row = sheet.getRow(rowNum);
			int cellType = row.getCell(cellNum).getCellType();	//获取单元格数据类型 进行switch匹配
			switch (cellType){
			case 0 :	//选择数值类型获取方法
				cellNumValue = row.getCell(cellNum).getNumericCellValue();
				System.out.print(cellNumValue);
				break;
			case 1 :	//选择字符串类型获取方法
				cellStrValue = row.getCell(cellNum).getStringCellValue();
				System.out.print(cellStrValue);
				break;
			default:
				break;
			}
		}
	}
	
	
}
