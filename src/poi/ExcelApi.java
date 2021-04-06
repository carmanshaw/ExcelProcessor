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
                //��ȡ��һ��sheet
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
	 * *��ȡ��������
	 * @return
	 */
	public List<String> getTitles() {
        //��ȡ��һ��
        Row row = sheet.getRow(0);
        //��ȡ�������
        int cols = row.getPhysicalNumberOfCells();
        
        lstTitles.clear();
        for(int i=0;i<cols;i++) {
        	String cellData = (String) getCellFormatValue(row.getCell(i));
        	lstTitles.add(i, cellData);
        }
        return lstTitles;
	}
	
	/**
	 * *����ѡ�е���title
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
	 * *��������������
	 */
	public void appendCont(String mainkey,Cell srcCell) {
		Row row = null;
	    //��ȡ�������
	    int rows = sheet.getPhysicalNumberOfRows();
	    
	    //�ҵ���Ӧ����
	    int i=0;
	    for(;i<rows;i++) {
	    	row = sheet.getRow(i);
    		String mkey = (String)getCellFormatValue(row.getCell(0));
    		if(mainkey.equals(mkey)) {
    			break;
    		}
	    }
	    //�ҵ����
	    if(i < rows) {
	    	//��ȡ�������
		  	int cols = row.getPhysicalNumberOfCells();
		  		
		    Cell distCell = row.createCell(cols);
			setCellFormatValue(distCell,srcCell);
	    }else {
	    	//û���ҵ�����������һ��
	    	Row endrow = sheet.createRow(rows);
	    	Cell keycell = endrow.createCell(0);
	    	keycell.setCellValue(mainkey);
	    	Cell contcell = endrow.createCell(lstTitles.size());
			setCellFormatValue(contcell,srcCell);
	    }
	    
	}
	
	/**
	 * �Ƚϱ��Cell��ʽ�������Ƿ�һ��
	 */
	private String compareError = "";
	public boolean compare(String mainkey,Cell srcCell) {
		Row row = null;
	    //��ȡ�������
	    int rows = sheet.getPhysicalNumberOfRows();
	    
	    //�ҵ���Ӧ����
	    int i=0;
	    for(;i<rows;i++) {
	    	row = sheet.getRow(i);
    		String mkey = (String)getCellFormatValue(row.getCell(0));
    		if(mainkey.equals(mkey)) {
    			break;
    		}
	    }
	    
	    compareError = "�Ҳ��� [" + mainkey + "] ";
	    //�ҵ����
	    if(i < rows) {
	    	Cell tmp = row.getCell(selectedId);
	    	compareError += "[" +  lstTitles.get(selectedId) + "]";
	    	if(tmp != null && srcCell != null) {
	    		if(tmp.getCellType() != srcCell.getCellType()) {
		    		compareError += " ��ʽ����! ";
		    	}
		    	
		    	if(!compare(srcCell,row.getCell(selectedId))) {
		    		compareError += " ���ݴ���! ";
		    	}else {
		    		return true;
		    	}
		    	
//		    	setFontStyle(tmp, Font.COLOR_RED);
	    	}else {
	    		compareError += " ���ݴ���! ";
	    	}
	    	
	    }else {
	    	compareError += " ������! ";
	    }
	    return false;
	}
	
	/**
	 * *���compareһ��ʹ�ã����ڻ�ȡ��������
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
            //�ж�cell�Ƿ�Ϊ���ڸ�ʽ
            if(DateUtil.isCellDateFormatted(src)){
                //ת��Ϊ���ڸ�ʽYYYY-mm-dd
                if(dist.getDateCellValue() == src.getDateCellValue()) {
                	return true;
                }
            }else{
                //����
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
		case Cell.CELL_TYPE_BLANK: // ��ֵ
			return true;
		case Cell.CELL_TYPE_ERROR: // ����
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
            //�ж�cell����
            switch(cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:{
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            }
            case Cell.CELL_TYPE_FORMULA:{
                //�ж�cell�Ƿ�Ϊ���ڸ�ʽ
                if(DateUtil.isCellDateFormatted(cell)){
                    //ת��Ϊ���ڸ�ʽYYYY-mm-dd
                    cellValue = cell.getDateCellValue();
                }else{
                    //����
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
			case Cell.CELL_TYPE_BLANK: // ��ֵ
				System.out.println("��ֵ");
				break;
			case Cell.CELL_TYPE_ERROR: // ����
				System.out.println(" ����");
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
        //�ж�cell����
        switch(src.getCellType()){
        case Cell.CELL_TYPE_NUMERIC:{
            dist.setCellValue(src.getNumericCellValue());
            break;
        }
        case Cell.CELL_TYPE_FORMULA:{
            //�ж�cell�Ƿ�Ϊ���ڸ�ʽ
            if(DateUtil.isCellDateFormatted(src)){
                //ת��Ϊ���ڸ�ʽYYYY-mm-dd
                dist.setCellValue(src.getDateCellValue());
            }else{
                //����
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
		case Cell.CELL_TYPE_BLANK: // ��ֵ
			System.out.println("��ֵ");
			break;
		case Cell.CELL_TYPE_ERROR: // ����
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
	 * ���������ʽ ��ҪΪ��ɫ
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
