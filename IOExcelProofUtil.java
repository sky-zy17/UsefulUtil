package com.enovell.gnet.util.resource;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

public class IOExcelProofUtil {

	private static String path = IOExcelProofUtil.class.getResource("").getPath();
	private static String getRootPath(String path){
		int temp=0;
		for(int i=0;i<path.length();i++){
			if(path.substring(i, i+8).equals("GNetwork")){
				temp=i;
				i=path.length();
			}
		}
		path=path.substring(0, temp);
		return path;
	}
	private static String rootPath=getRootPath(path);
	/**
	 * 校验sheet页
	 * @param sheet
	 * @param sheetNum
	 * @param proofCells
	 * @param sheetStartNum
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static boolean isSheetImport(Sheet sheet,int sheetNum,Map<String,int[]> proofCells,int sheetStartNum){
		List<String[]> errorList=new ArrayList<String[]>();//用来保存校对之后的错误信息
		String templateFilePath = rootPath +"GNetwork/resource/ImportErrorShow.xlsx";//错误校对表的模板
		String importErrorFilePath = rootPath+"GNetwork/downfile/ImportErrorShow.xlsx";//导入校对错误信息的输出文件
		
		Map<String,int[]> proofCellNum = null;
		if(proofCells != null){
			proofCellNum = proofCells;
		}
		
		try {
			File importErrorFile = new File(importErrorFilePath);
			if(!importErrorFile.exists()){
				importErrorFile.createNewFile();
			}
			FileOutputStream fos = new FileOutputStream(importErrorFile);
			Workbook wb = WorkbookFactory.create(new File(templateFilePath));
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
			cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			cellStyle.setBorderTop(CellStyle.BORDER_THIN);
			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
			cellStyle.setBorderRight(CellStyle.BORDER_THIN);
			errorList.clear();
			if(sheet == null){//本sheet页数据为空
				String[] errorArray = getErrorMessage(sheetNum,-1,-1,0);
				errorList.add(errorArray);
			}
			int sheetEndNum = sheet.getLastRowNum();
			if(sheetStartNum > sheetEndNum){//本sheet页数据为空
				String[] errorArray = getErrorMessage(sheetNum,-1,-1,0);
				errorList.add(errorArray);
			}else{
				for(int currentRow = sheetStartNum;currentRow <= sheetEndNum;){
					Object[] objs = checkRow(sheet,sheetNum,currentRow,proofCellNum);
					errorList.addAll((List<String[]>)objs[0]);
					currentRow = currentRow + (Integer)objs[1];
				}
			}
			//如果没有错误，关闭流，返回真
			if(errorList.size()==0){
				fos.close();
				return true;
			}
			Sheet s= wb.getSheetAt(0);//建立输入文件的一个sheet页
			Row r=s.createRow(0);//建立第一行
			//建立输出文件表的表头
			Cell cell1=r.createCell(0);
			cell1.setCellValue("sheet页");
			Cell cell2=r.createCell(1);
			cell2.setCellValue("行数");
			Cell cell3=r.createCell(2);
			cell3.setCellValue("列数");
			Cell cell4=r.createCell(3);
			cell4.setCellValue("具体错误情况");
			//写入校对之后的错误信息到输出文件中
			for (int j=0;j<errorList.size();j++) {
				Row row=s.createRow(j+1);
				Cell c1=row.createCell(0);
				c1.setCellValue(errorList.get(j)[0]);
				Cell c2=row.createCell(1);
				c2.setCellValue(errorList.get(j)[1]);
				Cell c3=row.createCell(2);
				c3.setCellValue(errorList.get(j)[2]);
				Cell c4=row.createCell(3);
				c4.setCellValue(errorList.get(j)[3]);
			}
			wb.write(fos);
			fos.close();
			
			return false;
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}

		return true;
	}
	/**
	 * 校验一行数据
	 * @param sheet
	 * @param sheetNum
	 * @param rowNum
	 * @param proofCells
	 * @return
	 */
	private static Object[] checkRow(Sheet sheet, int sheetNum, int rowNum, Map<String,int[]> proofCells){
		List<String[]> errorMessage = new ArrayList<String[]>();
		//实际上检查了多少行
		int betweenNum = 0;
		//不能为空的列
		int[] noNullCells = {};
		//是合并单元格的列
		int[] isMergedRegion = {};
		//数据只能为数字
		int[] isNumCells = {};
		
		if(proofCells != null){
			//得到数据不能为空的列
			if(proofCells.containsKey("notNull")){
				noNullCells = proofCells.get("notNull");
			}
			//得到是合并单元格的列
			if(proofCells.containsKey("isMerged")){
				isMergedRegion = proofCells.get("isMerged");
			}
			//得到只能为数字的列
			if(proofCells.containsKey("isNum")){
				isNumCells = proofCells.get("isNum");
			}
			//在此处可以继续添加条件
		}
		//=============================================
		if(proofCells.containsKey("isMerged")){
			betweenNum = isVirtualMergedRegion(sheet, rowNum, isMergedRegion);
			for(int colNum : noNullCells){
//				if(isMergedRegion(sheet, rowNum, colNum)){
//					if(sheet.getRow(rowNum) == null || sheet.getRow(rowNum).getCell(colNum) == null || StringUtils.isEmpty(getMergedRegionValue(sheet, rowNum, colNum))){
//						errorMessage.add(getErrorMessage(sheetNum, rowNum, colNum, 1));
//					}
//				}else{
					if(sheet.getRow(rowNum) == null || sheet.getRow(rowNum).getCell(colNum) == null || StringUtils.isEmpty(getCellValue(sheet.getRow(rowNum).getCell(colNum)))){
						errorMessage.add(getErrorMessage(sheetNum, rowNum, colNum, 1));
					}
//				}
			}
		}else{
			for(int colNum : noNullCells){
				if(isMergedRegion(sheet, rowNum, colNum)){
					betweenNum = getRowNum(sheet, rowNum, colNum);
					if(sheet.getRow(rowNum) == null || sheet.getRow(rowNum).getCell(colNum) == null || StringUtils.isEmpty(getMergedRegionValue(sheet, rowNum, colNum))){
						errorMessage.add(getErrorMessage(sheetNum, rowNum, colNum, 1));
					}
				}else{
					betweenNum = 1;
					if(sheet.getRow(rowNum) == null || sheet.getRow(rowNum).getCell(colNum) == null || StringUtils.isEmpty(getCellValue(sheet.getRow(rowNum).getCell(colNum)))){
						errorMessage.add(getErrorMessage(sheetNum, rowNum, colNum, 1));
					}
				}
			}
		}
		
		Object[] objs =  {errorMessage,betweenNum};
		return objs;
	}
	/**
	 * 判断是否是合并单元格
	 * @param sheet
	 * @param rowNum
	 * @param colNum
	 * @return
	 */
	private static boolean isMergedRegion(Sheet sheet, int rowNum, int colNum){
		int sheetMergeCount = sheet.getNumMergedRegions();
		for(int i = 0 ; i < sheetMergeCount ; i++ ){
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();
			
			if(rowNum >= firstRow && rowNum <= lastRow){
				if(colNum >= firstColumn && colNum <= lastColumn){
					return true ;
				}
			}
		}
		return false ;
	}
	/**
	 * 获得合并单元格的值
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	private static String getMergedRegionValue(Sheet sheet ,int row , int column){
		int sheetMergeCount = sheet.getNumMergedRegions();
		
		for(int i = 0 ; i < sheetMergeCount ; i++){
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();
			
			if(row >= firstRow && row <= lastRow){
					
				if(column >= firstColumn && column <= lastColumn){
					Row fRow = sheet.getRow(firstRow);
					Cell fCell = fRow.getCell(firstColumn);
					
					return getCellValue(fCell) ;
				}
			}
		}
		
		return null ;
	}
	/**
	 * 获得单元格的值
	 * @param cell
	 * @return
	 */
	private static String getCellValue(Cell cell) {
		String cellValue = "";
		if (cell != null) {
			switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:// 字符串类型
					cellValue = cell.toString().trim();
					if (cellValue.trim().equals("") || cellValue.trim().length() <= 0)
						cellValue = "";
					break;
				case Cell.CELL_TYPE_NUMERIC: // 数值类型
					if (DateUtil.isCellDateFormatted(cell)) {
						SimpleDateFormat simpleDateFormat = new SimpleDateFormat(
								"yyyy-MM-dd HH:mm:ss");
						java.util.Date theDate = cell.getDateCellValue();
						cellValue = simpleDateFormat.format(theDate);
					} else {
						HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
						cellValue = dataFormatter.formatCellValue(cell);
						if(cellValue.matches("^[K,k][0-9]+[+]$")){//处理自定义格式的公里标，因为自定义格式解析为数字
							cellValue = cellValue.substring(cellValue.indexOf("K")+1, cellValue.indexOf("+"));
							int cellValueInt = Integer.parseInt(cellValue);
							cellValue = "K"+cellValueInt/1000+"+"+cellValueInt%1000;
						}
//						cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
					}
					break;
				case Cell.CELL_TYPE_FORMULA: // 公式
					cellValue = "";
					break;
				case Cell.CELL_TYPE_BLANK:
					cellValue = "";
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					cellValue = String.valueOf(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_ERROR:
					break;
				default:
					break;
			}
		}
		return cellValue;
	}
	/**
	 * 根据错误类型获取对应单元格的错误信息
	 * @param sheetNum
	 * @param rowNum
	 * @param colNum
	 * @param errorType
	 * @return
	 */
	private static String[] getErrorMessage(int sheetNum, int rowNum, int colNum, int errorType){
		String[] errorMessage = new String[4];
		if(sheetNum == -1){
			errorMessage[0] = "";
		}else{
			errorMessage[0] = ""+(sheetNum+1);
		}
		if(rowNum == -1){
			errorMessage[1] = "";
		}else{
			errorMessage[1] = ""+(rowNum+1);
		}
		if(colNum == -1){
			errorMessage[2] = "";
		}else{
			errorMessage[2] = ""+(colNum+1);
		}
		switch(errorType){
			case 0:
				errorMessage[3] = "本sheet页数据为空";
				break;
			case 1:
				errorMessage[3] = "数据不能为空";
				break;
			//在此处继续添加错误类型
		}
		return errorMessage;
	}
	/**
	 * 获取合并单元格的最后一行的行号
	 * @param sheet
	 * @param rowNum
	 * @param colNum
	 * @return
	 */
/*	private static int getlastRowNum(Sheet sheet , int rowNum , int colNum){
		int sheetMergeCount = sheet.getNumMergedRegions();
		
		for(int i = 0 ; i < sheetMergeCount ; i++ ){
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();
			
			if(rowNum >= firstRow && rowNum <= lastRow){
				if(colNum >= firstColumn && colNum <= lastColumn){
					
					return lastRow;
				}
			}
		}
		
		return 0;
	}*/
	/**
	 * 获得合并单元格的行数
	 * @param sheet
	 * @param rowNum
	 * @param colNum
	 * @return
	 */
	private static int getRowNum(Sheet sheet , int rowNum , int colNum){
		int sheetMergeCount = sheet.getNumMergedRegions();
		
		for(int i = 0 ; i < sheetMergeCount ; i++ ){
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();
			
			if(rowNum >= firstRow && rowNum <= lastRow){
				if(colNum >= firstColumn && colNum <= lastColumn){
					
					return lastRow - firstRow;
				}
			}
		}
		
		return 0;
	}
	/**
	 * 利用反射机制实现简单的类的导入(无合并单元格、无关联其它对象，只有基本数据类型和String类型属性)
	 * @param sheet
	 * @param sheetStartNum
	 * @param tClass 传入T.class
	 * @return
	 */
	public static <T> List<T> simpleSheetImport(Sheet sheet, int sheetStartNum, Class<T> tClass){//此方法尚未用过
		Row row = null;
		List<T> list = new ArrayList<T>();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		try {
			for(int i = sheetStartNum;i < sheet.getLastRowNum();i++){
				row = sheet.getRow(i);
				T t = tClass.newInstance();
				Field[] fields = t.getClass().getDeclaredFields();
				for(int j = 0;j < row.getLastCellNum();j++){
					fields[j].setAccessible(true);
					if(fields[j].getType().isAssignableFrom(String.class)){//属性类型是String类型
						fields[j].set(t, getCellValue(row.getCell(j)));
					}else if(fields[j].getType().isAssignableFrom(Integer.class)){//属性类型是Integer类型
						fields[j].set(t, Integer.valueOf(getCellValue(row.getCell(j))));
					}else if(fields[j].getType().isAssignableFrom(Date.class)){//属性类型是Date类型
						Date date = sdf.parse(getCellValue(row.getCell(j)));
						fields[j].set(t, date);
					}
				}
				list.add(t);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}
	/**
	 * 对已经封装为对象(与其它表相关联)的数据进行二次验证
	 * @param sheetNum
	 * @param errors
	 */
	public static void emportErrorMessage(int sheetNum, Map<Integer,String> errors){
		String templateFilePath = rootPath +"GNetwork/resource/ImportErrorShow.xlsx";//错误校对表的模板
		String importErrorFilePath = rootPath+"GNetwork/downfile/ImportErrorShow.xlsx";//导入校对错误信息的输出文件
		try {
			File importErrorFile = new File(importErrorFilePath);
			if(!importErrorFile.exists()){
				importErrorFile.createNewFile();
			}
			FileOutputStream fos = new FileOutputStream(importErrorFile);
			Workbook wb = WorkbookFactory.create(new File(templateFilePath));
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
			cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			cellStyle.setBorderTop(CellStyle.BORDER_THIN);
			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
			cellStyle.setBorderRight(CellStyle.BORDER_THIN);
			Sheet s= wb.getSheetAt(0);//建立输入文件的一个sheet页
			Row r=s.createRow(0);//建立第一行
			//建立输出文件表的表头
			Cell cell1=r.createCell(0);
			cell1.setCellValue("sheet页");
			Cell cell2=r.createCell(1);
			cell2.setCellValue("行数");
			Cell cell3=r.createCell(2);
			cell3.setCellValue("列数");
			Cell cell4=r.createCell(3);
			cell4.setCellValue("具体错误情况");
			//写入校对之后的错误信息到输出文件中
			int curNum = 0;
			for(Integer i : errors.keySet()){
				curNum++;
				Row row = s.createRow(curNum);
				Cell c1 = row.createCell(0);
				c1.setCellValue(sheetNum+1);
				Cell c2 = row.createCell(1);
				c2.setCellValue(i);
				Cell c3 = row.createCell(2);
				c3.setCellValue("");
				Cell c4 = row.createCell(3);
				c4.setCellValue(errors.get(i));
			}
			
			wb.write(fos);
			fos.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}

	}
	/**
	 * 判断是否是合并单元格
	 * @param sheet
	 * @param rowNum
	 * @param isMergedRegion 合并单元格的行数
	 * @return
	 */
	private static int isVirtualMergedRegion(Sheet sheet, int rowNum, int[] isMergedRegion){
		int lastRowNum = rowNum;
		int stop = 0;
		Row row = sheet.getRow(rowNum);
		String[] str = new String[isMergedRegion.length];
		String[] newStr = new String[isMergedRegion.length]; 
		
		for(int j = 0;j <= isMergedRegion.length-1;j++){
			str[j] = getCellValue(row.getCell(isMergedRegion[j]));
		}
		while(true){
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int j = 0;j <= isMergedRegion.length-1;j++){
				newStr[j] = getCellValue(row.getCell(isMergedRegion[j]));
			}
			for(int i = 0;i <= isMergedRegion.length-1;i++){
				if(!newStr[i].equals(str[i]) && !newStr[i].trim().equals("")){
					stop = 1;
					break;
				}
			}
			if(stop == 1){
				break;
			}
		}

		return rowNum - lastRowNum;
	}
}
