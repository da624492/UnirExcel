package co.gov.sdm.unirExcel.servicio;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UtilExcel {
	public static String safeString(Object o) {
		if (o == null) {
			return null;
		}
		
		return o.toString();
	}	
	
	/**
	 * Retorna el valor de una celda
	 * */
	public static Object getCellValue(Cell c) {
		if (c == null) {
			return null;
		}
		CellType ct = c.getCellType();		
		CellStyle cs = c.getCellStyle();
		//System.out.println("ct name " + ct.name());
		//System.out.println("cs hidden " + cs.getHidden());
		//System.out.println("cs locked " + cs.getLocked());
		if (ct.equals(CellType.NUMERIC)) {
			Double val = c.getNumericCellValue();
			if ( Math.floor(val) == val.doubleValue() ) {
				int x = val.intValue();
				return x;
			}
			
			return val;
		}
		
		if (ct.equals(CellType.STRING)) {
			String s =c.getStringCellValue();
			if (s == null || s.trim() == "") {
				s = null;
			}
			return s;
		}
		
		return null;
	}	
	
	/**
	 * Copia la fila dada en la hoja dada en el indice dado
	 * */
	public static void copiarFila(Row r, Sheet shSalida, int indice) {
		if (r == null) {
			return;
		}
		
		Row rowSalida = shSalida.createRow(indice);
		
		int nc = 0;
		Iterator<Cell> itCell = r.cellIterator();
		while (itCell.hasNext()) {
			Cell c = itCell.next();
			Object valorCelda = getCellValue(c);
			if (valorCelda != null) {
				Cell cellSalida = rowSalida.createCell(nc);
				cellSalida.setCellValue(valorCelda.toString());
			}
			
			nc++;
		}
	}
	
	/**
	 * Crea una celda con el valor dado.
	 * */
	public static void celdaString(Row r, int posicion, String texto) {
		r.createCell(posicion).setCellValue(texto);
	}
	
	public static CellStyle estilosNegrilla(SXSSFWorkbook wb) {
		CellStyle estiloTitulos = wb.createCellStyle();              
        Font fuenteTitulo = wb.createFont();
        fuenteTitulo.setBold(true);
        estiloTitulos.setFont(fuenteTitulo);  
        return estiloTitulos;
	}
	
	
	public static String[] nombresHojas(File archivo) throws IOException {
		List<String> nombres = new ArrayList<String>();
		FileInputStream fis = new FileInputStream(archivo);
	    XSSFWorkbook wb = new XSSFWorkbook(fis);
	    int ns = wb.getNumberOfSheets();
	    for (int i = 0; i < ns; i++) {
	    	XSSFSheet sheet = wb.getSheetAt(i);
	    	nombres.add(sheet.getSheetName());
	    }
	    
	    return nombres.toArray(new String[] {});	    
	}
}