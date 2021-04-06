package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelApi {
	
	private Workbook wb = null;
	private Sheet sheet = null;
	private List<String> lstTitles = new ArrayList<String>();
	private int selectedId = -1;
	private String excelPath;
	
	public boolean loadFile(String filePath) {
		
		excelPath = filePath;
		
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                wb = new HSSFWorkbook(is);
                //获取第一个sheet
                sheet = wb.getSheetAt(0);
                return true;
            }else if(".xlsx".equals(extString)){
                wb = new XSSFWorkbook(is);
                sheet = wb.getSheetAt(0);
                return true;
            }else{
                wb = null;
            }
            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        return false;
	}
	
	public String getFilePath() {
		return excelPath;
	}
	/**
	 * *获取标题内容
	 * @return
	 */
	public List<String> getTitles() {
        //获取第一行
        Row row = sheet.getRow(0);
        //获取最大列数
        int cols = row.getPhysicalNumberOfCells();
        
        lstTitles.clear();
        for(int i=0;i<cols;i++) {
        	String cellData = (String) getCellFormatValue(row.getCell(i));
        	lstTitles.add(i, cellData);
        }
        return lstTitles;
	}
	
	/**
	 * *设置选中的列title
	 * @param cont
	 */
	public void setSelectedIndex(int index) {
		selectedId = index;
	}
	
	public int getSelectedIndex() {
		return selectedId;
	}
	
	public int getRows() {
		return sheet.getPhysicalNumberOfRows();
	}
	
	/**
	 * *在最后列添加内容
	 */
	public void appendCont(String mainkey,Cell srcCell) {
		Row row = null;
	    //获取最大行数
	    int rows = sheet.getPhysicalNumberOfRows();
	    
	    //找到对应的行
	    int i=0;
	    for(;i<rows;i++) {
	    	row = sheet.getRow(i);
    		String mkey = (String)getCellFormatValue(row.getCell(0));
    		if(mainkey.equals(mkey)) {
    			break;
    		}
	    }
	    //找到结果
	    if(i < rows) {
	    	//获取最大列数
		  	int cols = row.getPhysicalNumberOfCells();
		  		
		    Cell distCell = row.createCell(cols);
			setCellFormatValue(distCell,srcCell);
	    }else {
	    	//没有找到则在最后添加一行
	    	Row endrow = sheet.createRow(rows);
	    	Cell keycell = endrow.createCell(0);
	    	keycell.setCellValue(mainkey);
	    	Cell contcell = endrow.createCell(lstTitles.size());
			setCellFormatValue(contcell,srcCell);
	    }
	    
	}
	
	/**
	 * 比较表格Cell格式、内容是否一致
	 */
	private String compareError = "";
	public boolean compare(String mainkey,Cell srcCell) {
		Row row = null;
	    //获取最大行数
	    int rows = sheet.getPhysicalNumberOfRows();
	    
	    //找到对应的行
	    int i=0;
	    for(;i<rows;i++) {
	    	row = sheet.getRow(i);
    		String mkey = (String)getCellFormatValue(row.getCell(0));
    		if(mainkey.equals(mkey)) {
    			break;
    		}
	    }
	    
	    compareError = "右侧表格 [" + mainkey + "] ";
	    //找到结果
	    if(i < rows) {
	    	Cell tmp = row.getCell(selectedId);
	    	compareError += "[" +  lstTitles.get(selectedId) + "]";
	    	if(tmp != null && srcCell != null) {
	    		if(tmp.getCellType() != srcCell.getCellType()) {
		    		compareError += " 格式错误! ";
		    	}
		    	
		    	if(!compare(srcCell,row.getCell(selectedId))) {
		    		compareError += " 内容错误! ";
		    	}else {
		    		return true;
		    	}
		    	
//		    	setFontStyle(tmp, Font.COLOR_RED);
	    	}else {
	    		compareError += " 内容错误! ";
	    	}
	    	
	    }else {
	    	compareError += " 不存在! ";
	    }
	    return false;
	}
	
	/**
	 * *配合compare一起使用，用于获取错误内容
	 * @return
	 */
	public String getCompareError() {
		return compareError;
	}
	
	private boolean compare(Cell src,Cell dist) {
		
		if(src.getCellType() != dist.getCellType()) {
			return false;
		}
		
		switch(src.getCellType()){
        case Cell.CELL_TYPE_NUMERIC:{
            if(dist.getNumericCellValue() == src.getNumericCellValue()) {
            	return true;
            }
            break;
        }
        case Cell.CELL_TYPE_FORMULA:{
            //判断cell是否为日期格式
            if(DateUtil.isCellDateFormatted(src)){
                //转换为日期格式YYYY-mm-dd
                if(dist.getDateCellValue() == src.getDateCellValue()) {
                	return true;
                }
            }else{
                //数字
                if(dist.getNumericCellValue() == src.getNumericCellValue()) {
                	return true;
                }
            }
            break;
        }
        case Cell.CELL_TYPE_STRING:{
            if(dist.getRichStringCellValue().toString().equals(src.getRichStringCellValue().toString())) {
            	return true;
            }
            break;
        }
        case Cell.CELL_TYPE_BOOLEAN: // Boolean
			if(dist.getBooleanCellValue() == src.getBooleanCellValue()) {
            	return true;
            }
			break;
		case Cell.CELL_TYPE_BLANK: // 空值
			return true;
		case Cell.CELL_TYPE_ERROR: // 故障
			return false;
			
        default:break;
		}
		return false;
	}
	
	public Row getRowLine(int rowid) {
		
		return sheet.getRow(rowid);
	}
	
	public Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:{
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            }
            case Cell.CELL_TYPE_FORMULA:{
                //判断cell是否为日期格式
                if(DateUtil.isCellDateFormatted(cell)){
                    //转换为日期格式YYYY-mm-dd
                    cellValue = cell.getDateCellValue();
                }else{
                    //数字
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            }
            case Cell.CELL_TYPE_STRING:{
                cellValue = cell.getRichStringCellValue().getString();
                break;
            }
            case Cell.CELL_TYPE_BOOLEAN: // Boolean
				System.out.println(cell.getBooleanCellValue() + "\t");
				break;
			case Cell.CELL_TYPE_BLANK: // 空值
				System.out.println("空值");
				break;
			case Cell.CELL_TYPE_ERROR: // 故障
				System.out.println(" 故障");
				break;
				
            default:
                cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }
	
	public void setCellFormatValue(Cell dist,Cell src){
        //判断cell类型
        switch(src.getCellType()){
        case Cell.CELL_TYPE_NUMERIC:{
            dist.setCellValue(src.getNumericCellValue());
            break;
        }
        case Cell.CELL_TYPE_FORMULA:{
            //判断cell是否为日期格式
            if(DateUtil.isCellDateFormatted(src)){
                //转换为日期格式YYYY-mm-dd
                dist.setCellValue(src.getDateCellValue());
            }else{
                //数字
                dist.setCellValue(src.getNumericCellValue());
            }
            break;
        }
        case Cell.CELL_TYPE_STRING:{
            dist.setCellValue(src.getRichStringCellValue());
            break;
        }
        case Cell.CELL_TYPE_BOOLEAN: // Boolean
			dist.setCellValue(src.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_BLANK: // 空值
			System.out.println("空值");
			break;
		case Cell.CELL_TYPE_ERROR: // 故障
			System.out.println("NA");
			dist.setCellValue("NA");
			break;
			
        default:break;
        }
    }
	
	public void save(String path) throws IOException {
		File file = new File(path);
		if (file.exists()) {
			file.delete();
		}
		file.createNewFile();
		OutputStream os = new FileOutputStream(file);
		wb.write(os);
		os.close();
	}
	
	/**
	 * 设置字体格式 主要为颜色
	 * @param cell
	 * @param colors
	 */
	public void setFontStyle(Cell cell,short colors) {
		CellStyle dateStyle = wb.createCellStyle();
 
		Font font = wb.createFont();
		font.setColor(colors);
		
		dateStyle.setFont(font);
		cell.setCellStyle(dateStyle);
	}
}
