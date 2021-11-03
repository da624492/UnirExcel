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
			
			int colArchDem = -1;
			int colArchDemSeg = -1;
			int colArchSen = -1;
			int colArchDesc = -1;
			int colArchSenEle = -1;
			int colArchCont = -1;
			int colArchContSeg = -1;
			
			// Itera los archivos en el directorio home
		    File dir = new File(carpetaEntrada);
		    File[] archs = dir.listFiles();
		    boolean primerExcelLeido = false;
		    
		    for (int a = 0; a < archs.length; a++) {
		    	File arch = archs[a];
		    	if (arch.getName().endsWith(".xlsx")) {
		    		if (!primerExcelLeido) {
		    			colArchDem = copiarTitulos(arch, 0, 6, wbSalida, shDemarcacion).getLastCellNum()-1;
		    			colArchDemSeg = copiarTitulos(arch, 1, 0, wbSalida, shDemarcacionSeg).getLastCellNum()-1;
		    			colArchSen = copiarTitulos(arch, 2, 6, wbSalida, shSenalizacion).getLastCellNum()-1;
		    			colArchDesc = copiarTitulos(arch, 3, 7, wbSalida, shDescripcionTablero).getLastCellNum()-1;
		    			colArchSenEle = copiarTitulos(arch, 4, 6, wbSalida, shSenElevada).getLastCellNum()-1;
		    			colArchCont = copiarTitulos(arch, 5, 6, wbSalida, shControl).getLastCellNum()-1;
		    			colArchContSeg = copiarTitulos(arch, 6, 0, wbSalida, shControlSegmento).getLastCellNum()-1;
		    			primerExcelLeido = true;
		    		}
		    		
		    		filasCompiladoDem = extraerFilas2021(arch, 0, 7, shDemarcacion, filasCompiladoDem, colArchDem);
		    		filasCompiladoDemSeg = extraerFilas2021(arch, 1, 1, shDemarcacionSeg, filasCompiladoDemSeg, colArchDemSeg);
		    		filasCompiladoSen = extraerFilas2021(arch, 2, 7, shSenalizacion, filasCompiladoSen, colArchSen);
		    		filasCompiladoDesc = extraerFilas2021(arch, 3, 8, shDescripcionTablero, filasCompiladoDesc, colArchDesc);
		    		filasCompiladoSenEle = extraerFilas2021(arch, 4, 7, shSenElevada, filasCompiladoSenEle, colArchSenEle);
		    		filasCompiladoCont = extraerFilas2021(arch, 5, 7, shControl, filasCompiladoCont, colArchCont);
		    		filasCompiladoContSeg = extraerFilas2021(arch, 6, 1, shControlSegmento, filasCompiladoContSeg, colArchContSeg);
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