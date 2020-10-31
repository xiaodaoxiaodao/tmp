/**
 * 
 */
package com.til.mis.modulemanager.util;

import java.io.IOException;  
import java.io.OutputStream;  
import java.lang.reflect.Field;  
import java.lang.reflect.InvocationTargetException;  
import java.lang.reflect.Method;  
import java.text.SimpleDateFormat;  
import java.util.Collection;  
import java.util.Date;  
import java.util.HashMap;
import java.util.Iterator;  
import java.util.regex.Matcher;  
import java.util.regex.Pattern;  

import javax.servlet.http.HttpSession;

import org.apache.commons.lang.StringUtils;  
import org.apache.poi.hssf.usermodel.HSSFRichTextString;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

import com.chinasofti.ro.bizframework.core.libs.StringUtil;

public class ExcelUtil {
    public final static String EXCEl_FILE_2007 = "2007";  
    public final static String EXCEL_FILE_2003 = "2003";  
    private final static String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";  
    
    private Workbook workbook;
    private Sheet sheet;
    private HashMap<Integer, Integer> widthToColMap = new HashMap<Integer, Integer>();
    
    private String pattern;
    private String version;
    private CellStyle style;
      
    public void setPattern(String pattern) {
		this.pattern = pattern;
	}
    
	public void setVersion(String version) {
		this.version = version;
	}

	public void setStyle(CellStyle style) {
		this.style = style;
	}

	/** 
     * 导出无头部标题行Excel <br> 
     * 时间格式默认：yyyy-MM-dd hh:mm:ss <br> 
     * @param <T>
     * @param sheetName 表格标题 
     * @param dataset 数据集合 
     * @param out 输出流 
     * @param version 2003 或者 2007，不传时默认生成2007版本 
     */  
    public <T> void exportExcel(String sheetName, String[] fieldNames,Collection<T> dataset, OutputStream out,String version) { 
    	this.createWorkbook(version);        
    	this.createSingleSheet(sheetName, null, fieldNames, dataset, out);  
    }  
  
    /** 
     * <p> 
     * 导出带有头部标题行的Excel <br> 
     * 时间格式默认：yyyy-MM-dd hh:mm:ss <br> 
     * </p> 
     *  
     * @param title 表格标题 
     * @param headers 头部标题集合 
     * @param fieldNames 对象属性集合 
     * @param dataset 数据集合 
     * @param out 输出流 
     * @param version 2003 或者 2007，不传时默认生成2007版本 
     */  
    public <T> void exportExcel(String sheetName,String[] headers, String[] fieldNames, Collection<T> dataset,  
            OutputStream out,String version) {  
    	this.createWorkbook(version);
        this.createSingleSheet(sheetName, headers, fieldNames, dataset, out);  
    }  
    
    public <T> void exportExcel(String sheetName,String[] headers, String[] fieldNames, Collection<T> dataset,  
            OutputStream out,String version,HttpSession session) {  
    	this.createWorkbook(version);
        this.createSingleSheet(sheetName, headers, fieldNames, dataset, out,session);  
    } 
    
    public <T> void exportExcel(String sheetName,String[] headers, String[] fieldNames, Collection<T> dataset,  
            OutputStream out,String version,SharedRedis sharedRedis) {  
    	this.createWorkbook(version);
        this.createSingleSheet(sheetName, headers, fieldNames, dataset, out,sharedRedis);  
    } 
    
    public void createWorkbook(String version){
    	this.setVersion(version);
    	if(StringUtils.isBlank(version) || EXCEl_FILE_2007.equals(version.trim())){  
        	workbook = new XSSFWorkbook();
    	}else{  
        	workbook = new HSSFWorkbook();  
        } 
    	this.createCellStyle();
    }
      
    private void createSheet(String sheetName){
    	if (StringUtils.isBlank(sheetName)) {
    		sheet = workbook.createSheet();
		}
    	else{
    		sheet = workbook.createSheet(sheetName);
        }
    	sheet.setDefaultColumnWidth(15); 
    	//sheet.setColumnWidth(0, 256*100+184);
    }
    
    private void createCellStyle(){
        // 生成一个样式  
        style = workbook.createCellStyle();  
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);    	
    }
    
	public void mapWidthToColumn(Integer width, int colIndex) {
		this.widthToColMap.put(width, colIndex);
	}
	
    /** 
     * <p> 
     * 通用Excel导出方法,利用反射机制遍历对象的所有字段，将数据写入Excel文件中 <br> 
     * 此版本生成2007以上版本的文件 (文件后缀：xlsx) 
     * </p> 
     * @param <T>
     *  
     * @param title 
     *            表格标题名 
     * @param headers 
     *            表格头部标题集合 
     * @param dataset 
     *            需要显示的数据集合,集合中一定要放置符合JavaBean风格的类的对象。此方法支持的 
     *            JavaBean属性的数据类型有基本数据类型及String,Date 
     * @param out 
     *            与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中 
     * @param pattern 
     *            如果有时间数据，设定输出格式。默认为"yyyy-MM-dd hh:mm:ss" 
     */  
    @SuppressWarnings("rawtypes")
	public <T> void createSingleSheet(String sheetName, String[] headers, String[] fieldNames, Collection dataset, OutputStream out) {  
        this.createSheet(sheetName);  
        this.createSheetData(headers, fieldNames, dataset);
        this.closeWorkbook(out);  
    }  
    
    public <T> void createSingleSheet(String sheetName, String[] headers, String[] fieldNames, Collection dataset, OutputStream out,HttpSession session) {  
        this.createSheet(sheetName);  
        this.createSheetData(headers, fieldNames, dataset);
        this.closeWorkbook(out,session);  
    } 
    
    public <T> void createSingleSheet(String sheetName, String[] headers, String[] fieldNames, Collection dataset, OutputStream out,SharedRedis sharedRedis) {  
        this.createSheet(sheetName);  
        this.createSheetData(headers, fieldNames, dataset);
        this.closeWorkbook(out,sharedRedis);  
    } 
    
    @SuppressWarnings({ "rawtypes"})
    public <T> void createOneSheet(String sheetName, String[] headers, String[] fieldNames, Collection dataset) {  
        this.createSheet(sheetName);  
        this.createSheetData(headers, fieldNames, dataset);
    }    
    
    private void createCellHeader(String[] headers){
        // 产生表格标题行  
        Row row = sheet.createRow(0);  
        Cell cellHeader;  
        for (int i = 0; i < headers.length; i++) {  
            cellHeader = row.createCell(i);  
            cellHeader.setCellStyle(style);  
            if(StringUtils.isBlank(version) || EXCEl_FILE_2007.equals(version.trim())){  
            	cellHeader.setCellValue(new XSSFRichTextString(headers[i]));  
            }else{
            	cellHeader.setCellValue(new HSSFRichTextString(headers[i]));
            }
        }      	
    }
    
    @SuppressWarnings({ "unchecked", "rawtypes", "unused" })  
    private<T> void createSheetRows(String[] fieldNames, Collection dataset){
        // 遍历集合数据，产生数据行  
        Iterator<T> it = dataset.iterator();  
        int index = 0;  
        T t;  
        Field[] fields;  
        //Field field;  
        RichTextString richString;  
        Pattern p = Pattern.compile("^//d+(//.//d+)?$");  
        Matcher matcher;  
        String fieldName;  
        String getMethodName;  
        Cell cell;  
        Class tCls;  
        Method getMethod;  
        Object value;  
        String textValue;  
        SimpleDateFormat sdf = new SimpleDateFormat(StringUtil.isNotBlank(this.pattern)? this.pattern : DEFAULT_DATE_FORMAT);  
        while (it.hasNext()) {  
            index++;  
            Row row = sheet.createRow(index);  
            t = (T) it.next();  
            // 利用反射，根据JavaBean属性的先后顺序，动态调用getXxx()方法得到属性值  
            //fields = t.getClass().getDeclaredFields();  
            for (int i = 0; i < fieldNames.length; i++) {  
                cell = row.createCell(i);  
                //cell.setCellStyle(style2);  
                //field = fields[i];  
                fieldName = fieldNames[i];  
                getMethodName = "get" + fieldName.substring(0, 1).toUpperCase()  
                        + fieldName.substring(1);  
                try {  
                    tCls = t.getClass();  
                    getMethod = tCls.getMethod(getMethodName, new Class[] {});  
                    value = getMethod.invoke(t, new Object[] {});  
                    // 判断值的类型后进行强制类型转换  
                    textValue = null;  
                    if (value instanceof Integer) {  
                        cell.setCellValue((Integer) value); 
                    } else if (value instanceof Float) {  
                        textValue = String.valueOf((Float) value);  
                        cell.setCellValue(textValue);  
                    } else if (value instanceof Double) {  
                        textValue = String.valueOf((Double) value); 
                        cell.setCellValue(textValue);  
                    } else if (value instanceof Long) {  
                        cell.setCellValue((Long) value);  
                    }  
                    if (value instanceof Boolean) {  
                        textValue = "是";  
                        if (!(Boolean) value) {  
                            textValue = "否";  
                        }  
                    } else if (value instanceof Date) {  
                        textValue = sdf.format((Date) value);  
                    } else {  
                        // 其它数据类型都当作字符串简单处理  
                        if (value != null) {  
                            textValue = value.toString();  
                        }  
                    }  
                    if (textValue != null) {  
                        matcher = p.matcher(textValue);  
                        if (matcher.matches()) {  
                            // 是数字当作double处理  
                            cell.setCellValue(Double.parseDouble(textValue));  
                        } else {  
                        	if(StringUtils.isBlank(version) || EXCEl_FILE_2007.equals(version.trim())){  
                        		richString = new XSSFRichTextString(textValue);  
                            }else{
                            	richString = new HSSFRichTextString(textValue);
                            }
                            cell.setCellValue(richString);  
                        }  
                    }  
                    
                } catch (SecurityException e) {  
                    e.printStackTrace();  
                } catch (NoSuchMethodException e) {  
                    e.printStackTrace();  
                } catch (IllegalArgumentException e) {  
                    e.printStackTrace();  
                } catch (IllegalAccessException e) {  
                    e.printStackTrace();  
                } catch (InvocationTargetException e) {  
                    e.printStackTrace();  
                } finally {  
                    // 清理资源  
                } 
                //sheet.autoSizeColumn((short)i);
            }  
        }  
    	
    }
    
    private void createSheetData(String[] headers, String[] fieldNames, Collection dataset) {
        // 设置这些样式 
        /*
        style.setFillForegroundColor((short)java.awt.Color.BLUE.getBlue());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);  
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);  
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);  */
        // 生成一个字体  
        /*
        XSSFFont font = workbook.createFont();  
        //font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);  
        font.setFontName("宋体");   
        font.setColor(new XSSFColor(java.awt.Color.BLACK));  
        font.setFontHeightInPoints((short) 11);  
        // 把字体应用到当前的样式  
        style.setFont(font);  
        // 生成并设置另一个样式  
        XSSFCellStyle style2 = workbook.createCellStyle();  
        style2.setFillForegroundColor(new XSSFColor(java.awt.Color.WHITE));  
        style2.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);  
        style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);  
        style2.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style2.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);  
        style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);  
        // 生成另一个字体  
        XSSFFont font2 = workbook.createFont();  
        font2.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);  
        // 把字体应用到当前的样式  
        style2.setFont(font2);  */
  
    	createCellHeader(headers);
    	createSheetRows(fieldNames, dataset);    	
    }
    
     public void closeWorkbook(OutputStream out){
        try {  
			workbook.write(out);
			out.flush();
			out.close();
        } catch (IOException e) {  
            e.printStackTrace();  
        }      	
    }
     
     public void closeWorkbook(OutputStream out,HttpSession session){
         try {  
 			workbook.write(out);
 			out.flush();
 			out.close();
 			session.setAttribute("status", 1);
         } catch (IOException e) {  
             e.printStackTrace();  
         }      	
     }
     
     public void closeWorkbook(OutputStream out,SharedRedis sharedRedis){
         try {  
 			workbook.write(out);
 			out.flush();
 			out.close();
 			sharedRedis.set("status", "1");
         } catch (IOException e) {  
             e.printStackTrace();  
         }      	
     }
     
}
