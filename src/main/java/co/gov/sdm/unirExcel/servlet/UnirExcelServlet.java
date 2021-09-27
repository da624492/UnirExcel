package co.gov.sdm.unirExcel.servlet;

import java.io.File;
import java.io.FileInputStream;     
import java.io.FileNotFoundException; 
import java.io.FileOutputStream;  
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import co.gov.sdm.unirExcel.servicio.UtilExcel2021;
import co.gov.sdm.unirExcel.servicio.UtilExcelPre2021;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Controla la navegacion de la aplicacion.
 * */
@WebServlet(name = "UnirExcelServlet", urlPatterns = {"/servletExcel"})
public class UnirExcelServlet extends HttpServlet {
	@Override
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {		
		String dirOrigen = request.getParameter("directorioEntrada");
		String dirSalida = request.getParameter("directorioSalida");
		
		UtilExcel2021.extraer2021Todos(dirOrigen, dirSalida);
		request.setAttribute("mensaje", "Exceles compilados");
		request.getRequestDispatcher("/paginas/index.jsp").forward(request, response);
	}
}

