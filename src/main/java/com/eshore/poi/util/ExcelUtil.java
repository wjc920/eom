package com.eshore.poi.util;

import java.io.File;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 对象集合导出excel工具类 Created by wangqian on 2016/10/28.
 * @version 1.0 
 */

public class ExcelUtil {

	/**
	 * 对象集合导出到Excel
	 * @param data 待导出List
	 * @param type 导出excel后缀格式 
	 *            ExcelUtil.XLSX_TYPE对应.xlsx
	 *            ExcelUtil.XLS_TYPE对应.xls
	 * @param sheetName （可选）sheet名称,不填则自动命名，填写多个名称,只取第一个名称
	 * @return Workbook
	 * @author wangqian
	 */
	public static <T> Workbook exportToWorkbook(List<T> data, int type,
			String... sheetName) throws Exception {
		Workbook workbook = null;
		if (data.isEmpty()) {
			throw new IllegalArgumentException("数据列表为空");
		}
		if (type == XLS_TYPE) {
			workbook = new HSSFWorkbook();
		} else {
			workbook = new XSSFWorkbook();
		}
		// 创建sheet
		Sheet sheet = null;
		if (sheetName.length > 0) {
			sheet = workbook.createSheet(sheetName[0]);
		} else {
			sheet = workbook.createSheet();
		}
		T temp = data.get(0);
		short row = 0;
		// 写入列名
		{
			Method[] methods = temp.getClass().getDeclaredMethods();
			Row title = sheet.createRow(row++);
			for (Method m : methods) {
				if (m.isAnnotationPresent(ExcelColAttr.class)) {
					ExcelColAttr excelColAttr = (ExcelColAttr) m.getAnnotation(ExcelColAttr.class);
					title.createCell(excelColAttr.colIndex()).setCellValue(excelColAttr.colName());
				}
			}
		}
		// 写入数据
		fillDataToSheet(sheet, data, row);
		return workbook;
	}

	/**
	 * 从excel导入到对象集合
	 * @param cl   待映射的对象类型
	 * @param filename 待导入的excel文件名称
	 * @param sheetname 待导入的excel文件的sheet名称
	 * @param beginRow 指定从哪一行开始导入，最小值为0
	 * @return 对象集合
	 * @author wangqian
	 */
	public static <T> List<T> importFromFile(Class<T> cl,String filename,String sheetname,int beginRow) throws Exception{
		if(beginRow < 0){
			throw new IllegalArgumentException("行数不能为负值");
		}

		Workbook workbook = WorkbookFactory.create(new File(filename));
		if(workbook == null){
			throw new NullPointerException("创建Excel工作簿失败");
		}

		return importFromWorkbook(cl,workbook,sheetname,beginRow);
	}

	/**
	 * 从excel导入到对象集合
	 * @param cl   待映射的对象类型
	 * @param workbook 待导入的excel workbook对象
	 * @param sheetname 待导入的excel文件的sheet名称
	 * @param beginRow 指定从哪一行开始导入，最小值为0
	 * @return 对象集合
	 * @author wangqian
	 */
	public static <T> List<T> importFromWorkbook(Class<T> cl,Workbook workbook,String sheetname,int beginRow) throws Exception{
		if(beginRow < 0){
			throw new IllegalArgumentException("行数不能为负值");
		}

		if(workbook == null){
			throw new NullPointerException("创建Excel工作簿失败");
		}

		Sheet sheet = workbook.getSheet(sheetname);
		if(sheet == null){
			throw new NullPointerException("表:"+sheetname+"不存在");
		}

		List<T> data = new ArrayList<>();

		for(int i=beginRow;i<=sheet.getLastRowNum();i++){
			Row r = sheet.getRow(i);

			if(r == null){
				throw new IllegalArgumentException("行号"+i+"错误");
			}

			T temp = cl.newInstance();
			Method[] methods = cl.getMethods();
			for(Method m:methods){
				if(m.isAnnotationPresent(ExcelColAttr.class)){
					ExcelColAttr excelColAttr = (ExcelColAttr)m.getAnnotation(ExcelColAttr.class);
					String setMethodName = "s"+m.getName().substring(1);
					int colIndex = excelColAttr.colIndex();
					Method setMethod = null;
					Cell cell = r.getCell(colIndex,Row.RETURN_BLANK_AS_NULL);

					if(cell != null){
						if(excelColAttr.dataType() == ExcelColAttr.DataType.DOUBLE){
							setMethod = cl.getMethod(setMethodName,Double.class);
							Double value = 0.0;

							if(cell.getCellType() != Cell.CELL_TYPE_NUMERIC
									&& cell.getCellType() != Cell.CELL_TYPE_STRING){
								throw new IllegalArgumentException("行"+i+"列"+colIndex+"格式错误，不是数字或常规格式");
							}

							if(cell.getCellType() == Cell.CELL_TYPE_STRING){
								value = Double.parseDouble(cell.getStringCellValue());
							}
							else{
								value = cell.getNumericCellValue();
							}

							value = Double.parseDouble(new DecimalFormat(
									excelColAttr.doublePattern()).format(value));
							setMethod.invoke(temp,value);
						}
						else if(excelColAttr.dataType() == ExcelColAttr.DataType.INT){
							setMethod = cl.getMethod(setMethodName,Integer.class);

							if(cell.getCellType() != Cell.CELL_TYPE_NUMERIC
									&& cell.getCellType() != Cell.CELL_TYPE_STRING){
								throw new IllegalArgumentException("行"+i+"列"+colIndex+"格式错误，不是数字或常规格式");
							}

							if(cell.getCellType() == Cell.CELL_TYPE_STRING){
								setMethod.invoke(temp,Integer.parseInt(cell.getStringCellValue()));
							}
							else{
								setMethod.invoke(temp, Integer.valueOf((int)cell.getNumericCellValue()));
							}
						}
						else if(excelColAttr.dataType() == ExcelColAttr.DataType.DATE){
							setMethod = cl.getMethod(setMethodName,Date.class);
							SimpleDateFormat s = new SimpleDateFormat(excelColAttr.datePattern());

							if(cell.getCellType() != Cell.CELL_TYPE_NUMERIC &&
									cell.getCellType() != Cell.CELL_TYPE_STRING ){
								throw new IllegalArgumentException("行"+i+"列"+colIndex+"格式错误，不是日期或常规格式");
							}

							Date date = null;

							if(cell.getCellType() == Cell.CELL_TYPE_STRING){
								date = s.parse(cell.getStringCellValue());
							}
							else{
								if(DateUtil.isCellDateFormatted(cell)){
									date = cell.getDateCellValue();
									date = s.parse(s.format(date));
								}
								else{
									throw new IllegalArgumentException("行"+i+"列"+colIndex+"格式错误，不是日期或常规格式");
								}
							}

							setMethod.invoke(temp,date);
						}
						else if(excelColAttr.dataType() == ExcelColAttr.DataType.STRING){
							setMethod = cl.getMethod(setMethodName,String.class);

							if(cell.getCellType() != Cell.CELL_TYPE_STRING){
								throw new IllegalArgumentException("行"+i+"列"+colIndex+"格式错误，不是字符串");
							}

							setMethod.invoke(temp,cell.getStringCellValue());
						}
					}
					else{
						throw new IllegalArgumentException("行"+i+"列"+colIndex+"无数据");
					}
				}
			}

			data.add(temp);
		}

		return  data;
	}



	/**
	 * 在已有数据的excel中增加数据
	 * 
	 * @param workbook excel对象
	 * @param sheetName sheet名称
	 * @param data 待导出的数据
	 * @param startRow 导出时的起始行
	 * @throws Exception
	 * @author wangqian
	 */
	public static <T> void appendToWorkbook(Workbook workbook,
			String sheetName, List<T> data, int startRow) throws Exception {
		if(startRow < 0){
			throw new IllegalArgumentException("行数不能为负值");
		}

		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet=workbook.createSheet(sheetName);
			if(sheet==null){
			throw new NullPointerException("未获取到表名为"+sheetName+"的表");
			}
		}

		fillDataToSheet(sheet, data, startRow);
	}



	private static <T> void fillDataToSheet(Sheet sheet, List<T> data, int row)
			throws Exception {
		// 写入数据
		for (T val : data) {
			Method[] methods = val.getClass().getDeclaredMethods();
			Row dataRow = sheet.createRow(row++);
			for (Method m : methods) {
				if (m.isAnnotationPresent(ExcelColAttr.class)) {
					ExcelColAttr excelColAttr = (ExcelColAttr) m.getAnnotation(ExcelColAttr.class);
					Cell cell = dataRow.createCell(excelColAttr.colIndex());
					if (excelColAttr.dataType() == ExcelColAttr.DataType.INT) {
						Integer i = (Integer) m.invoke(val);
						cell.setCellValue(i);
					} else if (excelColAttr.dataType() == ExcelColAttr.DataType.DOUBLE) {
						Double d = (Double) m.invoke(val);
						cell.setCellValue(Double.parseDouble(new DecimalFormat(excelColAttr.doublePattern()).format(d)));
					} else if (excelColAttr.dataType() == ExcelColAttr.DataType.STRING) {
						cell.setCellValue((String) m.invoke(val));
					} else if (excelColAttr.dataType() == ExcelColAttr.DataType.DATE) {
						Date d = (Date) m.invoke(val);
						cell.setCellValue(new SimpleDateFormat(excelColAttr.datePattern()).format(d));
						sheet.setColumnWidth(excelColAttr.colIndex(), 5000);
					} else {
						throw new IllegalArgumentException("Error Type");
					}
				}
			}
		}
	}

	/**
	 * @Description 可动态指定部分添加ExcelColAttr标签的方法列导出
	 * @param sheet
	 * @param data
	 * @param notExpIndexList 不导出列的列下标集合
	 * @param row
	 * @throws Exception
	 * @author wangjichao
	 * @date 2016年11月7日 下午12:12:11
	 */
	private static <T> void fillDataToSheet(Sheet sheet, List<T> data,
			List<Integer> notExpIndexList, int row, List<Method> methodList) throws Exception {
		// 写入数据
		for (T val : data) {
			Row dataRow = sheet.createRow(row++);
			int j = 0;
			for (Method m : methodList) {
				ExcelColAttr excelColAttr = (ExcelColAttr) m.getAnnotation(ExcelColAttr.class);
				if (notExpIndexList.contains(excelColAttr.colIndex())) {
					j++;
					continue;
				}
				Cell cell = dataRow.createCell(excelColAttr.colIndex() - j);
				if (excelColAttr.dataType() == ExcelColAttr.DataType.INT) {
					Integer i = (Integer) m.invoke(val);
					cell.setCellValue(i);
				} else if (excelColAttr.dataType() == ExcelColAttr.DataType.DOUBLE) {
					Double d = (Double) m.invoke(val);
					cell.setCellValue(Double.parseDouble(new DecimalFormat(excelColAttr.doublePattern()).format(d)));
				} else if (excelColAttr.dataType() == ExcelColAttr.DataType.STRING) {
					cell.setCellValue((String) m.invoke(val));
				} else if (excelColAttr.dataType() == ExcelColAttr.DataType.DATE) {
					Date d = (Date) m.invoke(val);
					cell.setCellValue(new SimpleDateFormat(excelColAttr.datePattern()).format(d));
					sheet.setColumnWidth(excelColAttr.colIndex(), 5000);
				} else {
					throw new IllegalArgumentException("Error Type");
				}
			}
		}
	}
	
	/**
	 * 创建workbook并导入数据，同时可控制某几列不导入workbook
	 * @param data 待导入数据集合
	 * @param type 导入excel后缀格式 
	 * 				ExcelExporter.XLSX_TYPE对应.xlsx
	 *            	ExcelExporter.XLS_TYPE对应.xls
	 * @param indexList 不导出列的列下标集合
	 * @param sheetName sheet名称
	 * @return
	 * @throws Exception
	 * @author wangjichao
	 * @date 2016年12月2日 上午9:20:35
	 */
	public static <T> Workbook exportToWorkbook(List<T> data, int type,
			List<Integer> indexList, String... sheetName) throws Exception{
		Workbook workbook = null;
		if (data.isEmpty()) {
			throw new IllegalArgumentException("Empty Data List");
		}
		if (type == XLS_TYPE) {
			workbook = new HSSFWorkbook();
		} else {
			workbook = new XSSFWorkbook();
		}
		return exportToWorkbook(workbook, data, type, indexList, sheetName);
	}
	
	/**
	 * 在已有workbook基础上导入数据，并可控制某几列不导入excel
	 * @param workbook 待添加数据的workbook
	 * @param data 待导入数据集合
	 * @param type 导入excel后缀格式 
	 * 				ExcelExporter.XLSX_TYPE对应.xlsx
	 *            	ExcelExporter.XLS_TYPE对应.xls
	 * @param notExpIndexList 不导出列的列下标集合
	 * @param sheetName sheet名称
	 * @return
	 * @throws Exception
	 * @author wangjichao
	 * @date 2016年11月29日 上午11:42:56
	 */
	public static <T> Workbook exportToWorkbook(Workbook workbook, List<T> data, int type,
			List<Integer> notExpIndexList, String... sheetName) throws Exception {
		
		// 创建sheet
		Sheet sheet = null;
		if (sheetName.length > 0) {
			sheet = workbook.createSheet(sheetName[0]);
		} else {
			sheet = workbook.createSheet();
		}
		short row = 0;
		List<Method> methodList = null;
		// 写入列名
		{
			Method[] methods = data.get(0).getClass().getDeclaredMethods();
			methodList = new ArrayList<>();
			for (Method m : methods) {
				if (m.isAnnotationPresent(ExcelColAttr.class)) {
					methodList.add(m);
				}
			}
			Collections.sort(methodList, new Comparator<Method>() {
				@Override
				public int compare(Method o1, Method o2) {
					ExcelColAttr ExcelColAttr1 = (ExcelColAttr) o1.getAnnotation(ExcelColAttr.class);
					ExcelColAttr ExcelColAttr2 = (ExcelColAttr) o2.getAnnotation(ExcelColAttr.class);
					return ExcelColAttr1.colIndex() - ExcelColAttr2.colIndex();
				}
			});
			Row title = sheet.createRow(row++);
			int j = 0;
			for (Method m : methodList) {
				ExcelColAttr ExcelColAttr = (ExcelColAttr) m.getAnnotation(ExcelColAttr.class);
				if (notExpIndexList.contains(ExcelColAttr.colIndex())) {
					j++;
					continue;
				}
				title.createCell(ExcelColAttr.colIndex() - j).setCellValue(ExcelColAttr.colName());
			}
		}
		// 写入数据
		fillDataToSheet(sheet, data, notExpIndexList, row, methodList);
		return workbook;
	}
	
	/**
	 * 创建个性化表头
	 * @param cellorArray 表头各个单元格（合并单元格、简单单元格）数组
	 * @param type 导入excel后缀格式 
	 * 			  ExcelExporter.XLSX_TYPE对应.xlsx
	 *            ExcelExporter.XLS_TYPE对应.xls
	 * @param sheetName sheet名称
	 * @return
	 * @author wangjichao 
	 * @date 2016年12月2日 上午8:54:24
	 */
	public static Workbook creatTitleToSheet(Cellor[][] cellorArray, int type, String... sheetName) {
		Workbook workbook = null;
		if (type == XLS_TYPE) {
			workbook = new HSSFWorkbook();
		} else {
			workbook = new XSSFWorkbook();
		}
		return creatTitleToSheet(workbook, cellorArray, type,sheetName);
	}
	/**
	 * 在已有workbook上创建个性化表头
	 * @param cellorArray 表头各个单元格（合并单元格、简单单元格）数组
	 * @param type 导入excel后缀格式 
	 * 			  ExcelExporter.XLSX_TYPE对应.xlsx
	 *            ExcelExporter.XLS_TYPE对应.xls
	 * @param sheetName sheet名称
	 * @return
	 * @author wangjichao
	 * @date 2016年11月29日 上午11:53:53
	 */
	public static Workbook creatTitleToSheet(Workbook workbook, Cellor[][] cellorArray, int type,String... sheetName) {
		
		// 创建sheet
		Sheet sheet = null;
		if (sheetName.length > 0) {
			sheet = workbook.createSheet(sheetName[0]);
		} else {
			sheet = workbook.createSheet();
		}
		for (int i = 0; i < cellorArray.length; i++) {
			Row title = sheet.createRow(i);
			for (int j = 0; j < cellorArray[i].length; j++) {
				Cellor c = cellorArray[i][j];
				if (c.isMerged()) {
					CellRangeAddress cra = new CellRangeAddress(
							c.getFirstRow(), c.getLastRow(), c.getFirstCol(),
							c.getLastCol());
					sheet.addMergedRegion(cra);
				}
				title.createCell(c.getFirstCol()).setCellValue(
						c.getValue().toString());
			}
		}
		return workbook;
	}

	public static final int XLSX_TYPE = 0;
	public static final int XLS_TYPE  = 1;
}
