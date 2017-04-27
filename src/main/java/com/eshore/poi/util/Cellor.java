package com.eshore.poi.util;

/**    
 * @Description 单元格、合并单元格类 
 * @author wangjichao
 * @date 2016年11月3日 下午7:42:16 
 * @version 1.0   
 */
public class Cellor{
	private boolean isMerged;//是否为合并单元格 0：单元格 1：合并单元格
	private Integer firstRow;//单元格开始行
	private Integer lastRow;//单元格结束行
	private Integer firstCol;//单元格开始列
	private Integer lastCol;//单元格结束列
	private String value;//单元格的值
	/**
	 * 生成合并单元格
	 * @param firstRow 单元格开始行
	 * @param lastRow 单元格结束行
	 * @param firstCol 单元格开始列
	 * @param lastCol 单元格结束列
	 * @param value 单元格的值
	 */
	public Cellor(Integer firstRow, Integer lastRow, Integer firstCol,
			Integer lastCol, String value) {
		super();
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
		this.value = value;
		this.isMerged = true;
	}
	/**
	 * 生成简单单元格
	 * @param firstRow 所在行
	 * @param firstCol 所在列
	 * @param value 单元格的值
	 */
	public Cellor(Integer firstRow, Integer firstCol,String value) {
		super();
		this.firstRow = firstRow;
		this.firstCol = firstCol;
		this.value = value;
		this.isMerged = false;
	}
	public boolean isMerged() {
		return isMerged;
	}
	public void setMerged(boolean isMerged) {
		this.isMerged = isMerged;
	}
	public Integer getFirstRow() {
		return firstRow;
	}
	public void setFirstRow(Integer firstRow) {
		this.firstRow = firstRow;
	}
	public Integer getLastRow() {
		return lastRow;
	}
	public void setLastRow(Integer lastRow) {
		this.lastRow = lastRow;
	}
	public Integer getFirstCol() {
		return firstCol;
	}
	public void setFirstCol(Integer firstCol) {
		this.firstCol = firstCol;
	}
	public Integer getLastCol() {
		return lastCol;
	}
	public void setLastCol(Integer lastCol) {
		this.lastCol = lastCol;
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}

}
