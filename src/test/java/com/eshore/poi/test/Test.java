package com.eshore.poi.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.eshore.poi.util.Cellor;
import com.eshore.poi.util.ExcelUtil;

public class Test {
	public static void main(String[] args) throws Exception {
		//生成一个excel 生成表路径D://test.xlsx
	/*	int startRow=testExportToWorkbook();
		System.out.println("1.生成一个excel 生成表路径D://test.xlsx");
		//在上一步生成的excel基础上增加数据 生成表路径D://test-append.xlsx
		testAppendToWorkbook(++startRow);
		System.out.println("2.在上一步生成的excel基础上增加数据 生成表路径D://test-append.xlsx");
		List<Integer> indexList=new ArrayList<>();
		Collections.addAll(indexList, 1,3);//第1,3列不导出
		//根据需求选择对象的某几个属性不导出 生成表路径D://test-notall.xlsx
		testExportToWorkbookDynamic(indexList);
		System.out.println("3.根据需求选择对象的某几个属性不导出 生成表路径D://test-notall.xlsx");
		testExportToWorkbookDynamic2();
		System.out.println("4.根据需求选择对象的某几个属性不导出 生成表路径D://test-notall2.xlsx");
		//个性化生成表头 生成表路径D://test-title.xlsx
		testTitleToWorkbook();
		System.out.println("5.个性化生成表头 生成表路径D://test-title.xlsx");*/
	/*	testTitleToWorkbook2();
		System.out.println("6.个性化生成表头 生成表路径D://test-title2.xlsx");*/
//		testImportExcel();
		/*System.out.println("7.导入数据存在null值情况表路径D://test-null.xlsx");//zhangmeng的测试用例
		testExportNullPoint();*/
		System.out.println("7.导出数据单元格存在null值情况表路径D://test.xlsx");
		testImportNullPoint();
	}
	public static Integer testExportToWorkbook() throws Exception {
	    List<SimpleGrid> tests = new ArrayList<SimpleGrid>();
	    tests.add(new SimpleGrid(0,"张三丰",new Date(),60.112212));
	    tests.add(new SimpleGrid(1,"张三丰",new Date(),60.56));
	    tests.add(new SimpleGrid(2,"张三丰",new Date(),6.233));
	    tests.add(new SimpleGrid(3,"张三丰",new Date(),0.23));
	    tests.add(new SimpleGrid(4,"李三丰",new Date(),0.9));
	    tests.add(new SimpleGrid(5,"王三丰",new Date(),120.9));
	    tests.add(new SimpleGrid(6,"秦三丰",new Date(),60.12));
	    tests.add(new SimpleGrid(7,"张三丰",new Date(),60.12121));
	    tests.add(new SimpleGrid(8,"张三丰",new Date(),60.333));
	    tests.add(new SimpleGrid(9,"张三丰",new Date(),60.912));
	    tests.add(new SimpleGrid(10,"张三丰",new Date(),60.9123));

	    Workbook workbook = ExcelUtil.exportToWorkbook(tests, ExcelUtil.XLSX_TYPE);
	    FileOutputStream stream = new FileOutputStream("D://test.xlsx");
	    workbook.write(stream);
	    stream.close();
	    return tests.size();
	} 
	public static void testAppendToWorkbook(Integer startRow) throws Exception {
	    FileInputStream test = new FileInputStream("D://test.xlsx");
	    Workbook workbook = WorkbookFactory.create(test);
	    List<SimpleGrid> tests = new ArrayList<SimpleGrid>();
	    tests.add(new SimpleGrid(10,"张三丰",new Date(),60.112212));
	    tests.add(new SimpleGrid(11,"张三丰",new Date(),60.56));
	    tests.add(new SimpleGrid(12,"张三丰",new Date(),6.233));
	    tests.add(new SimpleGrid(13,"张三丰",new Date(),0.23));
	    tests.add(new SimpleGrid(14,"李三丰",new Date(),0.9));
	    tests.add(new SimpleGrid(15,"王三丰",new Date(),120.9));
	    tests.add(new SimpleGrid(16,"秦三丰",new Date(),60.12));
	    tests.add(new SimpleGrid(17,"张三丰",new Date(),60.12121));
	    tests.add(new SimpleGrid(18,"张三丰",new Date(),60.333));
	    tests.add(new SimpleGrid(19,"张三丰",new Date(),60.912));
	    tests.add(new SimpleGrid(110,"张三丰",new Date(),60.9123));
	    ExcelUtil.appendToWorkbook(workbook, "sheet0", tests, startRow);
	    FileOutputStream stream = new FileOutputStream("D://test-append.xlsx");
	    workbook.write(stream);
	    stream.close();
	}
	
	public static void testExportToWorkbookDynamic(List<Integer> indexList) throws Exception {
		List<SimpleGrid> tests = new ArrayList<SimpleGrid>();
		tests.add(new SimpleGrid(0,"张三丰",new Date(),60.112212));
		tests.add(new SimpleGrid(1,"张三丰",new Date(),60.56));
		tests.add(new SimpleGrid(2,"张三丰",new Date(),6.233));
		tests.add(new SimpleGrid(3,"张三丰",new Date(),0.23));
		tests.add(new SimpleGrid(4,"李三丰",new Date(),0.9));
		tests.add(new SimpleGrid(5,"王三丰",new Date(),120.9));
		tests.add(new SimpleGrid(6,"秦三丰",new Date(),60.12));
		tests.add(new SimpleGrid(7,"张三丰",new Date(),60.12121));
		tests.add(new SimpleGrid(8,"张三丰",new Date(),60.333));
		tests.add(new SimpleGrid(9,"张三丰",new Date(),60.912));
		tests.add(new SimpleGrid(10,"张三丰",new Date(),60.9123));
		
		Workbook workbook = ExcelUtil.exportToWorkbook(tests, ExcelUtil.XLSX_TYPE, indexList);
		FileOutputStream stream = new FileOutputStream("D://test-notall.xlsx");
		workbook.write(stream);
		stream.close();
	} 
	public static void testTitleToWorkbook() throws IOException{

		Cellor[][] titleArray=new Cellor[][]{
				{
					new Cellor(0, 1, 0, 0, "节点性质"),
					new Cellor(0, 1, 1, 1, "节点名称"),
					new Cellor(0, 0, 2, 3, "商户数量"),
					new Cellor(0, 1, 4, 5, "合计")
				},
				{
					new Cellor(1, 2, "肉类"),
					new Cellor(1, 3, "菜类")
				}
		};
		Workbook workbook= ExcelUtil.creatTitleToSheet(titleArray, ExcelUtil.XLSX_TYPE);
		FileOutputStream stream = new FileOutputStream("D://test-title.xlsx");
		workbook.write(stream);
		stream.close();
	}

	public static void testImportExcel(){
		List<SimpleGrid> simpleGridList = new ArrayList<>();

		try {
			simpleGridList = ExcelUtil.importFromFile(SimpleGrid.class, "D://test.xlsx", "Sheet0", 1);
			for (SimpleGrid s:simpleGridList){
				System.out.println(s.toString());
			}
		}
		catch (Exception e){
			e.printStackTrace();
			System.out.println(e.getMessage());
		}
	}
	public static void testTitleToWorkbook2() throws IOException{

		Cellor[][] titleArray=new Cellor[][]{
			
				{
					new Cellor(0, 0, "肉类"),
					new Cellor(0, 1, "菜类")
				}
		};
		Workbook workbook=ExcelUtil.creatTitleToSheet(titleArray, ExcelUtil.XLSX_TYPE);
		ExcelUtil.creatTitleToSheet(workbook, titleArray, ExcelUtil.XLSX_TYPE, "第二个表头");
		FileOutputStream stream = new FileOutputStream("C://test-title2.xlsx");
		workbook.write(stream);
		stream.close();
	}
	
	public static void testExportToWorkbookDynamic2() throws Exception {
		List<SimpleGrid> tests = new ArrayList<SimpleGrid>();
		tests.add(new SimpleGrid(0,"张三丰",new Date(),60.112212));
		tests.add(new SimpleGrid(1,"张三丰",new Date(),60.56));
		tests.add(new SimpleGrid(2,"张三丰",new Date(),6.233));
		tests.add(new SimpleGrid(3,"张三丰",new Date(),0.23));
		tests.add(new SimpleGrid(4,"李三丰",new Date(),0.9));
		tests.add(new SimpleGrid(5,"王三丰",new Date(),120.9));
		tests.add(new SimpleGrid(6,"秦三丰",new Date(),60.12));
		tests.add(new SimpleGrid(7,"张三丰",new Date(),60.12121));
		tests.add(new SimpleGrid(8,"张三丰",new Date(),60.333));
		tests.add(new SimpleGrid(9,"张三丰",new Date(),60.912));
		tests.add(new SimpleGrid(10,"张三丰",new Date(),60.9123));
		List<Integer> indexList=new ArrayList<>();
		Collections.addAll(indexList, 1,3);
		Workbook workbook = ExcelUtil.exportToWorkbook(tests, ExcelUtil.XLSX_TYPE,indexList,"13列");
		indexList.clear();
		Collections.addAll(indexList, 0,2);
		ExcelUtil.exportToWorkbook(workbook, tests,	ExcelUtil.XLSX_TYPE, indexList, "24列");
		FileOutputStream stream = new FileOutputStream("D://test-notall2.xlsx");
		workbook.write(stream);
		stream.close();
	} 
	
	public static void testExportNullPoint() throws Exception{
		 List<SimpleGrid> tests = new ArrayList<SimpleGrid>();
		    tests.add(new SimpleGrid(0,"张三丰",new Date(),null));
		    tests.add(new SimpleGrid(1,"张三丰",new Date(),null));
		    tests.add(new SimpleGrid(2,"张三丰",new Date(),null));
		    tests.add(new SimpleGrid(3,"张三丰",new Date(),null));
		    tests.add(new SimpleGrid(4,"李三丰",new Date(),null));
		    Workbook workbook = ExcelUtil.exportToWorkbook(tests, ExcelUtil.XLSX_TYPE);
		    FileOutputStream stream = new FileOutputStream("D://test-null.xlsx");
		    workbook.write(stream);
		    stream.close();
	}
	
	public static void testImportNullPoint() throws Exception{
		List<SimpleGrid> simpleGridList = new ArrayList<>();

		try {
			simpleGridList = ExcelUtil.importFromFile(SimpleGrid.class, "D://test.xlsx", "Sheet0", 1);
			for (SimpleGrid s:simpleGridList){
				System.out.println(s.toString());
			}
		}
		catch (Exception e){
			e.printStackTrace();
			System.out.println(e.getMessage());
		}
	}
	
}
