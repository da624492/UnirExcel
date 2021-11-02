package co.gov.sdm.unirExcel.servicio;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class UtilExcelWilber extends UtilExcelSDM {
	public static void extraer(String carpetaEntrada, String carpetaSalida) throws IOException {
		FileOutputStream out = null;
		SXSSFWorkbook wbSalida = null;
		
		try {		
			wbSalida = new SXSSFWorkbook(100);
			Sheet shDemarcacion = wbSalida.createSheet("DEM_DEMARCACION");
			Sheet shDemarcacionSeg = wbSalida.createSheet("DEM_DEMARCACION_SEG");
			Sheet shSenalizacion = wbSalida.createSheet("SEN_SENALIZACION");
			Sheet shDescripcionTablero = wbSalida.createSheet("SEN_DESCRIPCION_TABLERO");
			Sheet shSenElevada = wbSalida.createSheet("SEN_ELEVADA");
			Sheet shControl = wbSalida.createSheet("SS_CONTROL");
			Sheet shControlSegmento = wbSalida.createSheet("S_CONTROL_SEGMENTO");
			
			int filasCompiladoDem = 1;
			int filasCompiladoDemSeg = 1;
			int filasCompiladoSen = 1;
			int filasCompiladoDesc = 1;
			int filasCompiladoSenEle = 1;
			int filasCompiladoCont = 1;
			int filasCompiladoContSeg = 1;			
			
			// Itera los archivos en el directorio home
		    File dir = new File(carpetaEntrada);
		    File[] archs = dir.listFiles();
		    for (int a = 0; a < archs.length; a++) {
		    	File arch = archs[a];
		    	if (arch.getName().endsWith(".xlsx")) {		    		
		    		filasCompiladoDem = extraerFilas2021(arch, 0, 7, shDemarcacion, filasCompiladoDem, 29);
		    		filasCompiladoDemSeg = extraerFilas2021(arch, 1, 1, shDemarcacionSeg, filasCompiladoDemSeg, 29);
		    		filasCompiladoSen = extraerFilas2021(arch, 2, 7, shSenalizacion, filasCompiladoSen, 29);
		    		filasCompiladoDesc = extraerFilas2021(arch, 3, 8, shDescripcionTablero, filasCompiladoDesc, 29);
		    		filasCompiladoSenEle = extraerFilas2021(arch, 4, 7, shSenElevada, filasCompiladoSenEle, 29);
		    		filasCompiladoCont = extraerFilas2021(arch, 5, 7, shControl, filasCompiladoCont, 29);
		    		filasCompiladoContSeg = extraerFilas2021(arch, 6, 1, shControlSegmento, filasCompiladoContSeg, 29);		    				    		
		    	}
		    }
		    
		    // Finaliza la creacion del archivo de salida		    
		    SimpleDateFormat sdf = new SimpleDateFormat("dd_MM_yyyy");		    
		    String nombreSalida = "COMPILADO_"+sdf.format(new Date()) + ".xlsx";
		    out = new FileOutputStream(carpetaSalida + nombreSalida);
		    wbSalida.write(out);
			wbSalida.dispose();
		} finally {
			if (out != null)
				out.close();
			if (wbSalida != null)
				wbSalida.close();
		}
	}
}
