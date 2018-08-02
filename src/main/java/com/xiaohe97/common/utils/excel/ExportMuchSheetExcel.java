package com.xiaohe97.common.utils.excel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.google.common.collect.Lists;
import com.xiaohe97.common.utils.excel.annotation.ExcelField;

public class ExportMuchSheetExcel {
	
	
	
	
	
	/**
	 * 构造函数
	 * @param title 表格标题，传“空值”，表示无标题
	 * @param cls 实体对象，通过annotation.ExportField获取标题
	 * @param type 导出类型（1:导出数据；2：导出模板）
	 * @param groups 导入分组
	 */
	public void ExportExcel(SXSSFWorkbook wb,String title,int sheetNum,String sheetTitle,Class<?> cls, int type,int... groups){
		/**
		 * 注解列表（Object[]{ ExcelField, Field/Method }）
		 */
		List<Object[]> annotationList = Lists.newArrayList();
		
		// Get annotation field 
		Field[] fs = cls.getDeclaredFields();
		for (Field f : fs){
			ExcelField ef = f.getAnnotation(ExcelField.class);
			if (ef != null && (ef.type()==0 || ef.type()==type)){
				if (groups!=null && groups.length>0){
					boolean inGroup = false;
					for (int g : groups){
						if (inGroup){
							break;
						}
						for (int efg : ef.groups()){
							if (g == efg){
								inGroup = true;
								annotationList.add(new Object[]{ef, f});
								break;
							}
						}
					}
				}else{
					annotationList.add(new Object[]{ef, f});
				}
			}
		}
		// Get annotation method
		Method[] ms = cls.getDeclaredMethods();
		for (Method m : ms){
			ExcelField ef = m.getAnnotation(ExcelField.class);
			if (ef != null && (ef.type()==0 || ef.type()==type)){
				if (groups!=null && groups.length>0){
					boolean inGroup = false;
					for (int g : groups){
						if (inGroup){
							break;
						}
						for (int efg : ef.groups()){
							if (g == efg){
								inGroup = true;
								annotationList.add(new Object[]{ef, m});
								break;
							}
						}
					}
				}else{
					annotationList.add(new Object[]{ef, m});
				}
			}
		}
		// Field sorting
		Collections.sort(annotationList, new Comparator<Object[]>() {
			public int compare(Object[] o1, Object[] o2) {
				return new Integer(((ExcelField)o1[0]).sort()).compareTo(
						new Integer(((ExcelField)o2[0]).sort()));
			};
		});
		// Initialize
		List<String> headerList = Lists.newArrayList();
		for (Object[] os : annotationList){
			String t = ((ExcelField)os[0]).title();
			// 如果是导出，则去掉注释
			if (type==1){
				String[] ss = StringUtils.split(t, "**", 2);
				if (ss.length==2){
					t = ss[0];
				}
			}
			headerList.add(t);
		}
		initialize(wb,sheetNum,sheetTitle,title, headerList);
	}
	
	/**   
     * 使用已定义的数据源方式设置一个数据验证   
     * @param formulaString   
     * @param naturalRowIndex   
     * @param naturalColumnIndex   
     * @return   
     */    
    public void getDataValidationByFormula(Sheet sheet,String[] formulaString,int naturalColIndex){    
    	//加载下拉列表内容      
    	DataValidationHelper dvHelper = sheet.getDataValidationHelper();
    	DataValidationConstraint constraint = dvHelper.createExplicitListConstraint(formulaString);
    	// DVConstraint constraint = DVConstraint.createExplicitListConstraint(formulaString);     
        //设置数据有效性加载在哪个单元格上。      
        //四个参数分别是：起始行、终止行、起始列、终止列      
        int firstRow = 2;    
        int lastRow = 1048576-1;    
        int firstCol = naturalColIndex;    
        int lastCol = naturalColIndex;    
       
        CellRangeAddressList regions = new CellRangeAddressList(firstRow,lastRow,firstCol,lastCol);   
        DataValidation validation = dvHelper.createValidation(constraint, regions);
        //数据有效性对象     
        //DataValidation data_validation_list = new HSSFDataValidation(regions,constraint);
        if (validation instanceof XSSFDataValidation) {
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
        } else {
            validation.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(validation); 
    }
	
	/**
	 * 初始化函数
	 * @param title 表格标题，传“空值”，表示无标题
	 * @param headerList 表头列表
	 */
	private void initialize(SXSSFWorkbook wb,int sheetNum,String sheetTitle,String title, List<String> headerList) {
		//this.wb = new SXSSFWorkbook(500);
		Sheet sheet = wb.createSheet("Export");
		wb.setSheetName(sheetNum, sheetTitle);
		Map<String, CellStyle> styles = createStyles(wb);
		// Create title
		int rownum = 0;
		if (StringUtils.isNotBlank(title)){
			Row titleRow = sheet.createRow(rownum++);
			titleRow.setHeightInPoints(30);
			Cell titleCell = titleRow.createCell(0);
			titleCell.setCellStyle(styles.get("title"));
			titleCell.setCellValue(title);
			sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
					titleRow.getRowNum(), titleRow.getRowNum(), headerList.size()-1));
		}
		// Create header
		if (headerList == null){
			throw new RuntimeException("headerList not null!");
		}
		Row headerRow = sheet.createRow(rownum++);
		headerRow.setHeightInPoints(16);
		for (int i = 0; i < headerList.size(); i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellStyle(styles.get("header"));
			String[] ss = StringUtils.split(headerList.get(i), "**", 2);
			if (ss.length==2){
				cell.setCellValue(ss[0]);
				Comment comment = sheet.createDrawingPatriarch().createCellComment(
						new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
				comment.setString(new XSSFRichTextString(ss[1]));
				cell.setCellComment(comment);
			}else{
				cell.setCellValue(headerList.get(i));
			}
			sheet.autoSizeColumn(i);
		}
		for (int i = 0; i < headerList.size(); i++) {  
			int colWidth = sheet.getColumnWidth(i)*2;
	        sheet.setColumnWidth(i, colWidth < 3000 ? 3000 : colWidth);  
		}
		//log.debug("Initialize success.");
	}
	
	/**
	 * 创建表格样式
	 * @param wb 工作薄对象
	 * @return 样式列表
	 */
	private Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
		
		CellStyle style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		Font titleFont = wb.createFont();
		titleFont.setFontName("Arial");
		titleFont.setFontHeightInPoints((short) 16);
		titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(titleFont);
		styles.put("title", style);

		style = wb.createCellStyle();
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		Font dataFont = wb.createFont();
		dataFont.setFontName("Arial");
		dataFont.setFontHeightInPoints((short) 10);
		style.setFont(dataFont);
		styles.put("data", style);
		
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
		style.setAlignment(CellStyle.ALIGN_LEFT);
		styles.put("data1", style);

		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
		style.setAlignment(CellStyle.ALIGN_CENTER);
		styles.put("data2", style);

		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
		style.setAlignment(CellStyle.ALIGN_RIGHT);
		styles.put("data3", style);
		
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		Font headerFont = wb.createFont();
		headerFont.setFontName("Arial");
		headerFont.setFontHeightInPoints((short) 10);
		headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		headerFont.setColor(IndexedColors.WHITE.getIndex());
		style.setFont(headerFont);
		styles.put("header", style);
		
		return styles;
	}
	
    /** 
     * @Title: exportExcel 
     * @Description: 导出Excel的方法 
     * @author: evan @ 2014-01-09  
     * @param workbook  
     * @param sheetNum (sheet的位置，0表示第一个表格中的第一个sheet) 
     * @param sheetTitle  （sheet的名称） 
     * @param headers    （表格的标题） 
     * @param result   （表格的数据） 
     * @param out  （输出流） 
     * @throws Exception 
     */  
    public void exportExcel(HSSFWorkbook workbook, int sheetNum,  
            String sheetTitle, String[] headers, List<List<String>> result,  
            OutputStream out) throws Exception {  
        // 生成一个表格  
        HSSFSheet sheet = workbook.createSheet();  
        workbook.setSheetName(sheetNum, sheetTitle);  
        // 设置表格默认列宽度为20个字节  
        sheet.setDefaultColumnWidth((short) 20);  
        // 生成一个样式  
        HSSFCellStyle style = workbook.createCellStyle();  
        // 设置这些样式  
        style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);  
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        // 生成一个字体  
        HSSFFont font = workbook.createFont();  
        font.setColor(HSSFColor.BLACK.index);  
        font.setFontHeightInPoints((short) 12);  
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
        // 把字体应用到当前的样式  
        style.setFont(font);  
  
        // 指定当单元格内容显示不下时自动换行  
        style.setWrapText(true);  
  
        // 产生表格标题行  
        HSSFRow row = sheet.createRow(0);  
        for (int i = 0; i < headers.length; i++) {  
            HSSFCell cell = row.createCell((short) i);  
          
            cell.setCellStyle(style);  
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);  
            cell.setCellValue(text.toString());  
        }  
        // 遍历集合数据，产生数据行  
        if (result != null) {  
            int index = 1;  
            for (List<String> m : result) {  
                row = sheet.createRow(index);  
                int cellIndex = 0;  
                for (String str : m) {  
                    HSSFCell cell = row.createCell((short) cellIndex);  
                    cell.setCellValue(str.toString());  
                    cellIndex++;  
                }  
                index++;  
            }  
        }  
    }  
    
    
    @SuppressWarnings("unchecked")  
    public static void main(String[] args) {  
        try {  
            OutputStream out = new FileOutputStream("D:\\test.xls");  
            List<List<String>> data = new ArrayList<List<String>>();  
            for (int i = 1; i < 5; i++) {  
                List rowData = new ArrayList();  
                rowData.add(String.valueOf(i));  
                rowData.add("东霖柏鸿");  
                data.add(rowData);  
            }  
            String[] headers = { "ID", "用户名" };  
            ExportMuchSheetExcel eeu = new ExportMuchSheetExcel();  
            HSSFWorkbook workbook = new HSSFWorkbook();  
            eeu.exportExcel(workbook, 0, "上海", headers, data, out);  
            eeu.exportExcel(workbook, 1, "深圳", headers, data, out);  
            eeu.exportExcel(workbook, 2, "广州", headers, data, out);  
            //原理就是将所有的数据一起写入，然后再关闭输入流。  
            workbook.write(out);  
            out.close();  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
}
