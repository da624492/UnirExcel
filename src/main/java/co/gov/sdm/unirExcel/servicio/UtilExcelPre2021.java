package co.gov.sdm.unirExcel.servicio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import co.gov.sdm.unirExcel.dominio.IntervencionDemarcacion2021;


public class UtilExcelPre2021 extends UtilExcel {		
	public static void unir() throws IOException {
		String home = "C:/excelesAUnir";
		String dirCopias = home +"/" + "copias/";
		String dirSalida = home +"/" + "salida/";
	    File[] archs;	    
		File arch;
		File dir;
		
		// crea las copias sin blanks de los exceles a unir
	    dir = new File(home);
	    archs = dir.listFiles();
	    for (int a = 0; a < archs.length; a++) {
	    	arch = archs[a];
	    	if (arch.getName().endsWith(".xlsx")) {
	    		copiarArchivoSinBlanks(arch, dirCopias);
	    	}
	    }
	    
	    // Parametros de interseñalizacion para exceles sin blanks
		ParametrosHoja[] phs = new ParametrosHoja[] {
			new ParametrosHoja("HORIZONTAL", 5, 3),			
			new ParametrosHoja("VERTICAL",  5, 2),
			new ParametrosHoja("DEM_DEMARCACION_CIV", 1, 0)/*,
			new ParametrosHoja("SEN_DESCRIPCION_TABLERO", 3, 8),
			new ParametrosHoja("SEN_ELEVADA", 4, 7),
			new ParametrosHoja("SS_CONTROL", 5, 7),
			new ParametrosHoja("S_CONTROL_SEGMENTO", 6, 1)*/
		};
		
		SXSSFWorkbook wbSalida = new SXSSFWorkbook(100);
		
		for (int i = 0; i < phs.length; i++) {			
			ParametrosHoja ph = phs[i];			
			int numFilasEncabezado = ph.getNumFilasEncabezado();
			int numFilasPieDePagina = ph.getNumFilasPieDePagina();
			String nombreHoja = ph.getNombre();
			
			Sheet shSalida = wbSalida.createSheet(ph.getNombre());
			int filasCopiadas = 0;
		    
		    //System.out.println("filas copiadas " + filasCopiadas);
		    
		    dir = new File(dirCopias);
		    archs = dir.listFiles();
		    for (int a = 0; a < archs.length; a++) {
		    	arch = archs[a];
		    	if (arch.getName().endsWith(".xlsx")) {
		    		filasCopiadas = agregarArchivo(wbSalida, shSalida, arch, filasCopiadas, numFilasEncabezado,	numFilasPieDePagina,
		    		nombreHoja);		    		   
		    		//System.out.println("filas copiadas " + filasCopiadas);
		    	}			    
		    }		         		  
		}
		
		FileOutputStream out = new FileOutputStream(dirSalida + "sxssf.xlsx");
	    wbSalida.write(out);
	    out.close();
	    wbSalida.close();
	    wbSalida.dispose();	    	   					
	}
	
		
	/**
	 * Imprime en la consola el contenido de la primera hoja de un archivo de excel. 
	 * */
	public static void excelAConsola(File archivo) throws IOException {
		FileInputStream fis = new FileInputStream(archivo);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		try {
			int ns = 1;
			for (int i = 0; i < ns; i++) {
			    XSSFSheet sheet = wb.getSheetAt(i);		    
		    
			    int filas = 0;
			    Iterator<Row> it = sheet.iterator();
			    while (it.hasNext()) {
			    	Row r = it.next();
	
			    	// imprime las celdas de una fila: ini
					int nc = 0;
					Iterator<Cell> itCell = r.cellIterator();
					while (itCell.hasNext()) {
						Cell c = itCell.next();
						Object valorCelda = getCellValue(c);
						System.out.print(valorCelda + " ");										
						nc++;
					}
			    	// imprime las celdas de una fila: fin		    	
			    	System.out.println("");
				    filas++;
			    }
			}
		} finally {
		    fis.close();
		    wb.close();	    			
		}
	}
	
	/**
	 * Copia el contenido de la primera pagina de un excel a un nuevo archivo
	 * */
	public static void copiarArchivo(File archivo, String carpetaSalida) throws IOException {
		FileInputStream fis = new FileInputStream(archivo);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		SXSSFWorkbook wbSalida = new SXSSFWorkbook(100);
		
		int ns = wb.getNumberOfSheets();
		for (int i = 0; i < ns; i++) {
		    XSSFSheet sheet = wb.getSheetAt(i);
		    Sheet shSalida = wbSalida.createSheet(sheet.getSheetName());
	    
		    int filas = 0;
		    // procesa todas las filas en la hoja
		    Iterator<Row> it = sheet.iterator();
		    while (it.hasNext()) {
		    	Row r = it.next();
		    	copiarFila(r, shSalida, filas);
			    filas++;
		    }
		}
	    
	    fis.close();
	    wb.close();
	    
	    String[] tokens = archivo.getName().split("\\.");
	    String nombreSalida = tokens[0] +"_copia." + tokens[1];
	    FileOutputStream out = new FileOutputStream(carpetaSalida + nombreSalida);
	    wbSalida.write(out);
	    out.close();	    
	    wbSalida.dispose();
	    wbSalida.close();
	}
	
	private static void copiarArchivoSinBlanks(File archivo, String carpetaSalida) throws IOException {
		FileInputStream fis = new FileInputStream(archivo);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		SXSSFWorkbook wbSalida = new SXSSFWorkbook(100);
		
		int ns = wb.getNumberOfSheets();
		for (int i = 0; i < ns; i++) {
		    XSSFSheet sheet = wb.getSheetAt(i);
		    Sheet shSalida = wbSalida.createSheet(sheet.getSheetName());
	    
		    int filas = 0;
		    // procesa todas las filas en la hoja
		    Iterator<Row> it = sheet.iterator();
		    while (it.hasNext()) {
		    	Row r = it.next();
		    			    	
		    	Iterator<Cell> itCell = r.cellIterator();
	    		boolean filaVacia = true;
	    		while (itCell.hasNext()) {
	    			Cell c = itCell.next();
	    			Object val =getCellValue(c);
	    			if (val != null) {
	    				filaVacia = false;	    				
	    				break;
	    			}			    		
	    		}
	    		
	    		if (!filaVacia) {
	    			copiarFila(r, shSalida, filas);		    	
	    			filas++;	    		
	    		}
		    }
		    
		    //Row rowSalida = shSalida.createRow(filas);
		    //Cell cellSalida = rowSalida.createCell(0);
		    //cellSalida.setCellValue("fin");			
		}
	    
	    fis.close();
	    wb.close();
	    
	    String[] tokens = archivo.getName().split("\\.");
	    String nombreSalida = tokens[0] +"_copia." + tokens[1];
	    FileOutputStream out = new FileOutputStream(carpetaSalida + nombreSalida);
	    wbSalida.write(out);
	    out.close();	    
	    wbSalida.dispose();
	    wbSalida.close();
	}		

	private static int agregarArchivo(SXSSFWorkbook wbSalida, Sheet shSalida, File archivo, int filasCopiadas, 
	int numFilasEncabezado, int numFilasPieDePagina, String nombreHoja) throws IOException {
		FileInputStream fis = new FileInputStream(archivo);
	    XSSFWorkbook wb = new XSSFWorkbook(fis);
	    XSSFSheet sheet = wb.getSheet(nombreHoja);
	    
	    int filas = 0;
	    int ultimaFila = sheet.getLastRowNum() - numFilasPieDePagina;
	    
	    // procesa todas las filas en la hoja
	    Iterator<Row> it = sheet.iterator();	    
	    while (it.hasNext()) {
	    	Row r = it.next();	
	    		    	
	    	int rowNum = r.getRowNum();
	    	
	    	if (rowNum >= numFilasEncabezado && rowNum <= ultimaFila) {
	    		Iterator<Cell> itCell = r.cellIterator();
	    		boolean filaVacia = true;
	    		while (itCell.hasNext()) {
	    			Cell c = itCell.next();
	    			Object val =getCellValue(c); 	    			
	    			if (val != null) {
	    				filaVacia = false;	    				
	    				break;
	    			}    				    				    		
	    		}
	    		
	    		if (!filaVacia) {
	    			copiarFila(r, shSalida, filasCopiadas);
		    		filasCopiadas++;
	    		} else {
	    			//System.out.println("vacia");    			
	    		}	    		
	    	}
	    	
	    	filas++;
	    }
	    
	    
	    //System.out.println(sheet.getFirstRowNum());
	    //System.out.println(sheet.getLastRowNum());
	    //System.out.println(sheet.getPhysicalNumberOfRows());
	   /* 
	    int filasTotales = filas;
	    int pfpdp = filasTotales - numFilasPieDePagina;
	    // esto da row null
	    for (int i = pfpdp; i < filasTotales; i++) {
	    	Row r = shSalida.getRow(pfpdp);
	    	shSalida.removeRow(r);
	    	System.out.println("despues de borrar");
	    	System.out.println(sheet.getLastRowNum());
	    }
	  */

	    /*
	     
	    // esto da row null
	    for (int i = pfpdp; i < filasTotales; i++) {
	    	Row r = shSalida.getRow(i);
	    	shSalida.removeRow(r);	    
	    }
	     
	    // esto da row null
	    for (int i = 0; i < numFilasPieDePagina; i++) {
	    	Row r = shSalida.getRow(pfpdp);
	    	shSalida.removeRow(r);
	    }*/
	    
	    fis.close();
	    wb.close();
	    return filasCopiadas;
	}		
}

/*
IntervencionDemarcacion2021 intDem = new IntervencionDemarcacion2021();
intDem.setInterno( safeString(getCellValue(r.getCell(0))) );		    	
intDem.setClaseMarca( safeString(getCellValue(r.getCell(1))) );
intDem.setCiv( safeString(getCellValue(r.getCell(2))) );
intDem.setEjeTipoVia( safeString(getCellValue(r.getCell(3))) );
intDem.setEjeValor( safeString(getCellValue(r.getCell(4))) );
intDem.setIniciaTipoVia( safeString(getCellValue(r.getCell(5))) );
intDem.setIniciaValor( safeString(getCellValue(r.getCell(6))) );
intDem.setTerminaTipoVia( safeString(getCellValue(r.getCell(7))) );
intDem.setTerminaValor( safeString(getCellValue(r.getCell(8))) );
intDem.setFase( safeString(getCellValue(r.getCell(9))) );		    		
intDem.setAccion( safeString(getCellValue(r.getCell(10))) );
intDem.setEstado( safeString(getCellValue(r.getCell(11))) );
intDem.setFechaFase( safeString(getCellValue(r.getCell(12))) );
intDem.setNumeroCuenta( safeString(getCellValue(r.getCell(13))) );
intDem.setTipoMedida( safeString(getCellValue(r.getCell(14))) );
intDem.setUnidadCantidad( safeString(getCellValue(r.getCell(15))) );
intDem.setPinturaCantidad( safeString(getCellValue(r.getCell(16))) );
intDem.setPinturaColor( safeString(getCellValue(r.getCell(17))) );
intDem.setImprimanteCantidad( safeString(getCellValue(r.getCell(18))) );
intDem.setImprimanteColor( safeString(getCellValue(r.getCell(19))) );		    		
intDem.setAntideslizanteCantidad( safeString(getCellValue(r.getCell(20))) );
intDem.setAntideslizanteInterno( safeString(getCellValue(r.getCell(21))) );
intDem.setGarantia( safeString(getCellValue(r.getCell(22))) );
intDem.setGarantiaFechaVencimiento( safeString(getCellValue(r.getCell(23))) );
intDem.setInternoItemPintura( safeString(getCellValue(r.getCell(24))) );
intDem.setInternoItemImprimante( safeString(getCellValue(r.getCell(25))) );
intDem.setInternoItemInstalacion( safeString(getCellValue(r.getCell(26))) );
intDem.setObservaciones( safeString(getCellValue(r.getCell(27))) );
*/