//package poi;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.io.OutputStream;
//import java.util.ArrayList;
//import java.util.LinkedHashMap;
//import java.util.List;
//import java.util.Map;
//import java.util.Map.Entry;
//
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.CreationHelper;
//import org.apache.poi.ss.usermodel.DateUtil;
//import org.apache.poi.ss.usermodel.Font;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFCellStyle;
//import org.apache.poi.xssf.usermodel.XSSFDataFormat;
//import org.apache.poi.xssf.usermodel.XSSFRichTextString;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class Testpoi {
//
//    public static void main(String[] args) {
//    	
//    	testWrite();
//    }
//    
//    public static void testWrite() {
//    	
//    	List<Export> list = new ArrayList<Export>();
//    	
//    	for(int i=0;i<10;i++) {
//    		Export e = new Export();
//	    	e.setNum(100);
//	    	e.setSjbm("carman");
//	    	list.add(e);
//    	}
//    	String filePath = "F:\\WorkSpace\\ExcelApp\\test1.xlsx";
//    	try {
//			writeXls(list,new File(filePath));
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//    }
//    
//    public static void writeXls(List<Export> exportList, File file) throws Exception {
//    	 
//		String[] options = { "条码", "批次号", "数量" };
//		XSSFWorkbook book = new XSSFWorkbook();
// 
//		CreationHelper createHelper = book.getCreationHelper();
// 
//		XSSFCellStyle style = book.createCellStyle();
//		XSSFCellStyle dateStyle = book.createCellStyle();
//		XSSFDataFormat format = book.createDataFormat();
//		style.setWrapText(true);
//		dateStyle.setWrapText(true);
// 
//		XSSFSheet sheet = book.createSheet("sheet");
// 
//		sheet.setColumnWidth(3, 13000);
//		sheet.setDefaultColumnWidth(20);
// 
//		XSSFRow firstRow = sheet.createRow(0);
//		XSSFCell[] firstCells = new XSSFCell[3];
// 
//		CellStyle styleBlue = book.createCellStyle(); // 样式对象
//		// 设置单元格的背景颜色为淡蓝色
//		styleBlue.setFillBackgroundColor(HSSFColor.GREEN.index);
// 
//		styleBlue.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直
//		styleBlue.setAlignment(CellStyle.ALIGN_CENTER);// 水平
//		styleBlue.setWrapText(true);// 指定当单元格内容显示不下时自动换行
// 
//		Font font = book.createFont();
//		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
//		font.setFontName("宋体");
//		font.setFontHeight((short) 280);
//		font.setColor(Font.COLOR_RED);
//		
//		style.setFont(font);
//		dateStyle.setFont(font);
//		dateStyle.setDataFormat(format.getFormat("yyyy-mm-dd"));
//		styleBlue.setFont(font);
// 
//		for (int j = 0; j < options.length; j++) {
//			firstCells[j] = firstRow.createCell(j);
//			firstCells[j].setCellStyle(styleBlue);
//			firstCells[j].setCellValue(new XSSFRichTextString(options[j]));
//		}
//		getExport(sheet, style, createHelper, exportList, dateStyle);
//		if (file.exists()) {
//			file.delete();
//		}
//		boolean bret = file.createNewFile();
//		System.out.println("createNewFile:" + bret);
//		OutputStream os = new FileOutputStream(file);
//		book.write(os);
//		os.close();
//	}
//    
//    public static void getExport(XSSFSheet sheet, XSSFCellStyle style, CreationHelper createHelper, List<Export> exportList,
//			XSSFCellStyle dateStyle) {
//		for (int i = 0; i < exportList.size(); i++) {
//			XSSFRow row = sheet.createRow(i + 1);
// 
//			Export export = exportList.get(i);
//			XSSFCell hotelId = row.createCell(0);
//			hotelId.setCellStyle(style);
//			XSSFCell hotelName = row.createCell(1);
//			hotelName.setCellStyle(dateStyle);
//			XSSFCell chargeCount = row.createCell(2);
//			chargeCount.setCellStyle(style);
// 
//			hotelId.setCellValue(export.getSjbm());
//			hotelName.setCellValue("2018-3-1");
//			chargeCount.setCellValue(export.getNum());
// 
//			// ta.append("写入excel开始,行数是" + (i + 1) + "\n");
//		}
// 
//	}
//
//}