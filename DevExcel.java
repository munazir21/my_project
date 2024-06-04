package in.co.greenwave;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import javax.faces.context.ExternalContext;
import javax.faces.context.FacesContext;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;


public class DevExcel {


	public static void setCell(Sheet sheet, int rowNum, XSSFCellStyle cellStyle, int colNum, String text,Integer colWidth) {


		Row row = sheet.getRow(rowNum);
		if(row == null) {
			row= sheet.createRow(rowNum);
		}

		Cell cell = row.createCell(colNum);
		/*try{

			Double d = Double.parseDouble(text);
			cell.setCellValue(d);
		}catch(Exception e){
			cell.setCellValue(text);
		}*/
		cell.setCellValue(text);

		if(cellStyle==null)
			cellStyle  = defautStyle();
		cell.setCellStyle(cellStyle);
		if(colWidth == null)
			colWidth = 16*256;
		sheet.setColumnWidth(colNum, colWidth);
	}

	public static void setCell(Sheet sheet, int rowNum, XSSFCellStyle cellStyle, int colNum, Double text,Integer colWidth) {


		Row row = sheet.getRow(rowNum);
		if(row == null) {
			row= sheet.createRow(rowNum);
		}

		Cell cell = row.createCell(colNum);
		/*try{

			Double d = Double.parseDouble(text);
			cell.setCellValue(d);
		}catch(Exception e){
			cell.setCellValue(text);
		}*/
		cell.setCellValue(text);

		if(cellStyle==null)
			cellStyle  = defautStyle();
		cell.setCellStyle(cellStyle);
		if(colWidth == null)
			colWidth = 16*256;
		sheet.setColumnWidth(colNum, colWidth);
	}


	/*public static void download(SXSSFWorkbook wb,String fileName )
	{


		FacesContext fc = FacesContext.getCurrentInstance();
		ExternalContext ec = fc.getExternalContext();

		File file = new File(fileName);
		String attachment = "attachment;filename="+file.getName();
		FileOutputStream fileOut;
		FileInputStream fileIn = null;
		OutputStream output = null;

		try {
			fileOut = new FileOutputStream(file);
			wb.write(fileOut);
			ec.responseReset(); // Some JSF component library or some Filter might have set some headers in the buffer beforehand. We want to get rid of them, else it may collide.
			ec.setResponseContentType("application/vnd.ms-excel"); // Check http://www.iana.org/assignments/media-types for all types. Use if necessary ExternalContext#getMimeType() for auto-detection based on filename.
			ec.setResponseContentLength((int) file.length()); // Set it with the file size. This header is optional. It will work if it's omitted, but the download progress will be unknown.
			ec.setResponseHeader("Content-Disposition", attachment); // The Save As popup magic is done here. You can give it any file name you want, this only won't work in MSIE, it will use current request URL as file name instead.
			fileOut.flush();
			fileOut.close();
			//File file1 = new File("D:\\"+"INVReport on "+Calendar.getInstance().getTimeInMillis()+".xlsx");
			fileIn = new FileInputStream(file);

			output = ec.getResponseOutputStream();

			byte[] outputByte = new byte[4096];

			//copy binary contect to output stream
			while(fileIn.read(outputByte, 0, 4096) != -1)
			{
				output.write(outputByte, 0, 4096);
			}

			fileIn.close();
			output.flush();
			output.close();
			fc.responseComplete();
			file.delete();
		}
		catch(IOException e) {
			e.printStackTrace();
		}
		finally {
			try {
				fileIn.close();
				output.flush();
				output.close();
				fc.responseComplete();
				file.delete();
			}catch(Exception e) {

			}
		}

	}*/
	public static void download(SXSSFWorkbook wb,String fileName )
	{

        System.out.println("in download sercvice file name is"+fileName);
		FacesContext fc = FacesContext.getCurrentInstance();
		ExternalContext ec = fc.getExternalContext();

		File file = new File(fileName);
		String attachment = "attachment;filename="+file.getName();
		FileOutputStream fileOut;
		FileInputStream fileIn = null;
		OutputStream output = null;
		try {
			fileOut = new FileOutputStream(file);
			wb.write(fileOut);
			ec.responseReset(); // Some JSF component library or some Filter might have set some headers in the buffer beforehand. We want to get rid of them, else it may collide.
			ec.setResponseContentType("application/vnd.ms-excel"); // Check http://www.iana.org/assignments/media-types for all types. Use if necessary ExternalContext#getMimeType() for auto-detection based on filename.
			ec.setResponseContentLength((int) file.length()); // Set it with the file size. This header is optional. It will work if it's omitted, but the download progress will be unknown.
			ec.setResponseHeader("Content-Disposition", attachment); // The Save As popup magic is done here. You can give it any file name you want, this only won't work in MSIE, it will use current request URL as file name instead.
			fileOut.flush();
			fileOut.close();
			//File file1 = new File("D:\\"+"INVReport on "+Calendar.getInstance().getTimeInMillis()+".xlsx");
			fileIn = new FileInputStream(file);

			output = ec.getResponseOutputStream();

			byte[] outputByte = new byte[4096];

			//copy binary contect to output stream
			// while(fileIn.read(outputByte, 0, 4096) != -1)
			// {
			// output.write(outputByte, 0, 4096);
			// }
			int readBytes=0;
			//copy binary contect to output stream
			while((readBytes = fileIn.read(outputByte)) > 0)
			{
				try {
					output.write(outputByte, 0, readBytes);
				}catch(Exception e) {
					e.printStackTrace();
				}
			}

			// fileIn.close();
			// output.flush();
			// output.close();
			// fc.responseComplete();
			// file.delete();
		}
		catch(IOException e) {
			e.printStackTrace();
		}
		finally {
			try {
 
				output.flush();
				output.close();
				fc.responseComplete();
				file.delete();
			}catch(Exception e) {
				e.printStackTrace();
			}
		}

	}

	public static XSSFCellStyle defautStyle() {
		SXSSFWorkbook wb =  new SXSSFWorkbook();
		XSSFCellStyle cellStyle = (XSSFCellStyle) wb.createCellStyle();

		Font font = wb.createFont();
		font.setFontHeightInPoints((short)10);
		font.setBold(false);
		cellStyle.setFont(font);


		return cellStyle;
	}


	public static void autoSizeColumns(Workbook workbook) {
		int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			SXSSFSheet sheet = (SXSSFSheet) workbook.getSheetAt(i);
			sheet.trackAllColumnsForAutoSizing();
			if (sheet.getPhysicalNumberOfRows() > 0) {
				Row row = sheet.getRow(0);
				if(row!=null) {
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						int columnIndex = cell.getColumnIndex();

						sheet.autoSizeColumn(columnIndex);
					}
				}
			}
		}
	}

}
