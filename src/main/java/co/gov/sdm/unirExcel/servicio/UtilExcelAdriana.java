package co.gov.sdm.unirExcel.servicio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UtilExcelAdriana extends UtilExcelSDM {
	/**
	 * Itera los archivos en carpetaEntrada y crea un compilado en carpetaSalida
	 * */
	public static void extraer2021Todos(String carpetaEntrada, String carpetaSalida) throws IOException {		
		FileOutputStream out = null;
		SXSSFWorkbook wbSalida = null;
		
		try {		
			wbSalida = new SXSSFWorkbook(100);
			Sheet shDemarcacion = wbSalida.createSheet("DEMARCACION");
			Sheet shVertical = wbSalida.createSheet("VERTICAL");
			Sheet shCIV = wbSalida.createSheet("CIV_DEM");
			tituloDemarcacion(wbSalida, shDemarcacion);
			tituloVertical(wbSalida, shVertical);
			tituloCIV(wbSalida, shCIV);

			int filasCompiladoDem = 1;
			int filasCompiladoVert = 1;
			int filasCompiladoCIV = 1;
			
			// Itera los archivos en el directorio home
		    File dir = new File(carpetaEntrada);
		    File[] archs = dir.listFiles();
		    for (int a = 0; a < archs.length; a++) {
		    	File arch = archs[a];
		    	if (arch.getName().endsWith(".xlsx")) {
		    		filasCompiladoDem = extraerFilas2021(arch, 0, 9, shDemarcacion, filasCompiladoDem, 29);
		    		filasCompiladoVert = extraerFilas2021(arch, 1, 8, shVertical, filasCompiladoVert, 32);
		    		filasCompiladoCIV = extraerFilas2021(arch, 2, 2, shCIV, filasCompiladoCIV, 3);
		    	}
		    }
			
		    // Finaliza la creacion del archivo de salida
		    //probarValidacion(wbSalida);
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
	
	private static void probarValidacion( SXSSFWorkbook workbook) {        
		Sheet sheet = workbook.createSheet("Data Validation");
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		DataValidationConstraint dvConstraint = dvHelper.createExplicitListConstraint(new String[]{"13", "23", "33"});
		CellRangeAddressList addressList = new CellRangeAddressList(2, 2, 1, 1);            
		DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
		  // Note the check on the actual type of the DataValidation object.
		  // If it is an instance of the XSSFDataValidation class then the
		  // boolean value 'false' must be passed to the setSuppressDropDownArrow()
		  // method and an explicit call made to the setShowErrorBox() method.
		/*
		if(validation instanceof XSSFDataValidation) {
		    validation.setSuppressDropDownArrow(true);
		    validation.setShowErrorBox(true);
		}
		else {
		    // If the Datavalidation contains an instance of the HSSFDataValidation
		    // class then 'true' should be passed to the setSuppressDropDownArrow()
		    // method and the call to setShowErrorBox() is not necessary.
		    validation.setSuppressDropDownArrow(false);
		}*/
		sheet.addValidationData(validation);
	}
		
	/**
	 * Crea la fila de titulos de la hoja de CIV
	 * */
	private static void tituloCIV(SXSSFWorkbook wbSalida, Sheet shSalida) {
		Row rowSalida = shSalida.createRow(0);		        		
		rowSalida.setRowStyle(estilosNegrilla(wbSalida));
		celdaString(rowSalida, 0, "CONSECUTIVO");
		celdaString(rowSalida, 1, "INTERNO_DEMARCACION");
		celdaString(rowSalida, 2, "CIV");
		celdaString(rowSalida, 3, "ARCHIVO ORIGINAL");
	}

	/**
	 * Crea la fila de titulos de la hoja de vertical
	 * */
	private static void tituloVertical(SXSSFWorkbook wbSalida, Sheet shSalida) {
		Row rowSalida = shSalida.createRow(0);
		rowSalida.setRowStyle(estilosNegrilla(wbSalida));
		celdaString(rowSalida, 0, "No");
		celdaString(rowSalida, 1, "Interno");
		celdaString(rowSalida, 2, "TIPO_SENAL");
		celdaString(rowSalida, 3, "CLASE_SENAL");
		celdaString(rowSalida, 4, "TIPO_PEDESTAL");
		celdaString(rowSalida, 5, "CIV");
		celdaString(rowSalida, 6, "DIRECCION_EJE_TIPO");
		celdaString(rowSalida, 7, "DIRECCION_EVE_VALOR");
		celdaString(rowSalida, 8, "DIRECCION_INICIO_TIPO");
		celdaString(rowSalida, 9, "DIRECCION_INICIO_VALOR");
		celdaString(rowSalida, 10, "DIRECCION_TERMINA_TIPO");
		celdaString(rowSalida, 11, "DIRECCION_TERMINA_VALOR");
		celdaString(rowSalida, 12, "FASE");
		celdaString(rowSalida, 13, "ACCION");
		celdaString(rowSalida, 14, "ESTADO");
		celdaString(rowSalida, 15, "FECHA_FASE");
		celdaString(rowSalida, 16, "NUMERO_CUENTA");
		celdaString(rowSalida, 17, "TIPO_REFLECTIVO");
		celdaString(rowSalida, 18, "MATERIAL_TABLERO");
		celdaString(rowSalida, 19, "DIMENSIONES_ANCHO");
		celdaString(rowSalida, 20, "DIMENSIONES_ALTO");
		celdaString(rowSalida, 21, "CONTENIDO");
		celdaString(rowSalida, 22, "TIPO_FLECHA");		
		celdaString(rowSalida, 23, "ITEM_ANTIGRAFITI_INTERNO");
		celdaString(rowSalida, 24, "ITEM_ANTIGRAFITI_CANTIDAD");
		celdaString(rowSalida, 25, "ITEM_SUMINISTRO_INTERNO");
		celdaString(rowSalida, 26, "ITEM_SUMINISTRO_CANTIDAD");
		celdaString(rowSalida, 27, "ITEM_INSTALACION_INTERNO");
		celdaString(rowSalida, 28, "ITEM_INSTALACION_CANTIDAD");
		celdaString(rowSalida, 29, "ITEM_RET_REU_INTERNO");
		celdaString(rowSalida, 30, "ITEM_RET_REU_CANTIDAD");
		celdaString(rowSalida, 31, "OBSERVACIONES");
		celdaString(rowSalida, 32, "ARCHIVO ORIGINAL");
	}
	
	/**
	 * Crea la fila de titulos de la hoja de demarcacion
	 * */
	private static void tituloDemarcacion(SXSSFWorkbook wbSalida, Sheet shSalida) {
		Row rowSalida = shSalida.createRow(0);
		rowSalida.setRowStyle(estilosNegrilla(wbSalida));
		celdaString(rowSalida, 0, "No");
		celdaString(rowSalida, 1, "Interno");
		celdaString(rowSalida, 2, "CLASE_MARCA");
		celdaString(rowSalida, 3, "CIV");
		celdaString(rowSalida, 4, "TRAMO EJE TIPO");
		celdaString(rowSalida, 5, "TRAMO EJE VALOR");
		celdaString(rowSalida, 6, "TRAMO INICIA TIPO");
		celdaString(rowSalida, 7, "TRAMO INICIA VALOR");
		celdaString(rowSalida, 8, "TRAMO TERMINA TIPO");
		celdaString(rowSalida, 9, "TRAMO TERMINA VALOR");
		celdaString(rowSalida, 10, "FASE");
		celdaString(rowSalida, 11, "ACCION");
		celdaString(rowSalida, 12, "ESTADO");
		celdaString(rowSalida, 13, "FECHA_FASE");
		celdaString(rowSalida, 14, "NUMERO_CUENTA");
		celdaString(rowSalida, 15, "TIPO_MEDIDA");
		celdaString(rowSalida, 16, "UNIDAD CANTIDAD");
		celdaString(rowSalida, 17, "PINTURA CANTIDAD");
		celdaString(rowSalida, 18, "PINTURA COLOR");
		celdaString(rowSalida, 19, "IMPRIMANTE CANTIDAD");
		celdaString(rowSalida, 20, "IMPRIMANTE COLOR");
		celdaString(rowSalida, 21, "ANTIDESLIZANTE INTERNO");
		celdaString(rowSalida, 22, "ANTIDESLIZANTE CANTIDAD");
		celdaString(rowSalida, 23, "GARANTIA");
		celdaString(rowSalida, 24, "VENCIMIENTO GARANTIA");
		celdaString(rowSalida, 25, "INT. ITEM PINTURA");
		celdaString(rowSalida, 26, "INT. ITEM IMPRIMANTE");
		celdaString(rowSalida, 27, "INT. ITEM INSTALACION");
		celdaString(rowSalida, 28, "OBSERVACIONES");
		celdaString(rowSalida, 29, "ARCHIVO ORIGINAL");
	}		
}