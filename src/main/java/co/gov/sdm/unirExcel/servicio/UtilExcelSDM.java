package co.gov.sdm.unirExcel.servicio;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UtilExcelSDM extends UtilExcel {
	/**
	 * Extrae el contenido de la fila dada de la hoja dada de un archivo en formato 2021 y lo pega en shSalida
	 * @return La fila de titulos
	 * */
	public static Row copiarTitulos(File archivo, int hoja, int indFilaTitulos, SXSSFWorkbook wbSalida, Sheet shSalida) throws IOException {				
		FileInputStream fis = new FileInputStream(archivo);
		XSSFWorkbook wb = new XSSFWorkbook(fis);						
	    try {
			XSSFSheet sheet = wb.getSheetAt(hoja);		   
	    	Row r = sheet.getRow(indFilaTitulos);
	    	Row rt = copiarFila(r, shSalida, 0);
	    	rt.createCell( rt.getLastCellNum() ).setCellValue("ARCHIVO ORIGINAL");
	    	rt.setRowStyle(estilosNegrilla(wbSalida));
	    	return rt;
	    } finally {
	    	if (fis != null)
	    		fis.close();
	    	if (wb != null)
	    		wb.close();
	    }
	}
	
	/**
	 * Extrae el contenido de la hoja dada de un archivo en formato 2021 y lo pega en shSalida
	 * */
	public static int extraerFilas2021(File archivo, int hoja, int primeraFilaDeDatos, Sheet shSalida, int filasCompilado, int colNombArch) throws IOException {				
		FileInputStream fis = new FileInputStream(archivo);
		XSSFWorkbook wb = new XSSFWorkbook(fis);						
	    try {
			XSSFSheet sheet = wb.getSheetAt(hoja);		   
		    int filas = 0;
		    String nombArch = archivo.getName();
		    		
		    // procesa todas las filas en la hoja
		    Iterator<Row> it = sheet.iterator();
		    while (it.hasNext()) {
		    	Row r = it.next();
		    	if (filas >= primeraFilaDeDatos) {
		    		String num = safeString(getCellValue(r.getCell(0)));
		    		
		    		if (num == null || num.trim().length() == 0) {
		    			break;
		    		}
		    	
		    		Row rowSalida =copiarFila(r, shSalida, filasCompilado);
		    		rowSalida.createCell(colNombArch).setCellValue(nombArch);
		    		filasCompilado++;
		    	}
		    	
			    filas++;
		    }
		    
		    return filasCompilado;
	    } finally {
	    	if (fis != null)
	    		fis.close();
	    	if (wb != null)
	    		wb.close();
	    }
	}
}