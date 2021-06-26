package co.gov.sdm.unirExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

/**
 * Controla la navegacion de la aplicacion.
 * */
public class UnirExcelServlet extends HttpServlet {
	@Override
	protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
		super.doGet(req, resp);
		
		UtilExcel.unir();				
		//UtilExcel.escribir();
	}
}

class UtilExcel {		
	public static int unir() throws IOException {
		int filasCopiadas = 0;
		SXSSFWorkbook wbSalida = new SXSSFWorkbook(100);	    
	    Sheet shSalida = wbSalida.createSheet();
			    	    
	    System.out.println("filas copiadas " + filasCopiadas);
	    filasCopiadas = agregarArchivo(wbSalida, shSalida, "C:/datos/SH_SV_IDU_1558_2017_FICHA31.xlsx", filasCopiadas);
	    System.out.println("filas copiadas " + filasCopiadas);
	    filasCopiadas = agregarArchivo(wbSalida, shSalida, "C:/datos/SH_SV_IDU_1558_2017_FICHA32.xlsx", filasCopiadas);
	    System.out.println("filas copiadas " + filasCopiadas);
	    filasCopiadas = agregarArchivo(wbSalida, shSalida, "C:/datos/SH_SV_IDU_1558_2017_FICHA33.xlsx", filasCopiadas);
	    System.out.println("filas copiadas " + filasCopiadas);
	    
	    FileOutputStream out = new FileOutputStream("c:/datos/sxssf.xlsx");
	    wbSalida.write(out);
	    out.close();
	    wbSalida.dispose();
	    
	    return filasCopiadas;
	}
	
	public static Object getCellValue(Cell c) {
		CellType ct = c.getCellType();
		if (ct.equals(CellType.NUMERIC)) {
			return c.getNumericCellValue();
		}
		
		if (ct.equals(CellType.STRING)) {
			return c.getStringCellValue();
		}
		
		return null;
	}
	
	public static int agregarArchivo(SXSSFWorkbook wbSalida, Sheet shSalida, String rutaArchivo, int filasCopiadas) throws IOException {
		FileInputStream fis = new FileInputStream(rutaArchivo);
	    XSSFWorkbook wb = new XSSFWorkbook(fis);
	    XSSFSheet sheet = wb.getSheetAt(0);
	    Iterator<Row> it = sheet.iterator();
	    while (it.hasNext()) {
	    	Row r = it.next();	    	
	    	
	    	if (r.getRowNum() > 6) {
	    		Iterator<Cell> itCell = r.cellIterator();
	    		boolean filaVacia = true;
	    		while (itCell.hasNext()) {
	    			Cell c = itCell.next();
	    			if (getCellValue(c) != null) {
	    				filaVacia = false;
	    				break;
	    			}
	    		}	    		
	    		if (!filaVacia) {
		    		Row rowSalida = shSalida.createRow(filasCopiadas);
		    		
		    		int nc = 0;
		    		itCell = r.cellIterator();
		    		while (itCell.hasNext()) {
		    			Cell c = itCell.next();	    			
		    			if (c.getCellType() == CellType.STRING) {
		    				Cell cellSalida = rowSalida.createCell(nc);	    			
		    				cellSalida.setCellValue(c.getStringCellValue());	    			
		    			}
		    			nc++;
		    		}
		    		filasCopiadas++;
	    		}
	    	}
	    }	
	    
	    return filasCopiadas;
	}		
}