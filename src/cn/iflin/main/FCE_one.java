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
		String filePath = "F:\\Project\\16��Ŀ��\\��Ժ������γ�\\2015-2016ѧ��\\2015-2016-1���޿�.xlsx";
		FileInputStream excelFile = new FileInputStream(
				new File(filePath));
		XSSFWorkbook  workbook = new XSSFWorkbook(excelFile);
		GetCellNum(workbook);
	}
	
	/*
	 * ��ȡexcel�ļ���������������
	 */
	private static void GetCellNum(XSSFWorkbook  workbook) throws IOException{
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFRow row = sheet.getRow(2);
		int cellNum,cellAllNum = row.getLastCellNum();
		String cellValue=null;
		for(cellNum=0;cellNum<cellAllNum;cellNum++){
			cellValue = row.getCell(cellNum).getStringCellValue();	//��ȡ��Ԫ���ڵ�ֵ
			switch (cellValue){	//��ȡ����ֵ����ƥ�� �����ϵ���������OutExcel����
			case "��ѡ": 
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "�γ�����":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "��ʦ����":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "�γ�����":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			case "ѧ��":
				OutExcel(cellNum,cellValue);
				System.out.println(cellValue);
				break;
			default:
				break;
			}
		}
	}
	/*
	 * ��ȡexcel�ļ����������������е���������
	 * ���޸�
	 */
	@SuppressWarnings("resource")
	private static  void OutExcel(int cellNum,String cellFirstValue) throws IOException{
		/*��ȡ��Excel�ļ�*/
		String filePath = "F:\\Project\\16��Ŀ��\\��Ժ������γ�\\2015-2016ѧ��\\2015-2016-1���޿�.xlsx";
		FileInputStream excelFile = new FileInputStream(
				new File(filePath));
		XSSFWorkbook  workbook = new XSSFWorkbook(excelFile);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		XSSFRow row ;
		String cellStrValue = null;
		double cellNumValue;
		int rowNum,rowAllNum=sheet.getLastRowNum(); //ȡ����ֵ
		for(rowNum=3;rowNum<rowAllNum;rowNum++){
			row = sheet.getRow(rowNum);
			int cellType = row.getCell(cellNum).getCellType();	//��ȡ��Ԫ���������� ����switchƥ��
			switch (cellType){
			case 0 :	//ѡ����ֵ���ͻ�ȡ����
				cellNumValue = row.getCell(cellNum).getNumericCellValue();
				System.out.print(cellNumValue);
				break;
			case 1 :	//ѡ���ַ������ͻ�ȡ����
				cellStrValue = row.getCell(cellNum).getStringCellValue();
				System.out.print(cellStrValue);
				break;
			default:
				break;
			}
		}
	}
	
	
}
