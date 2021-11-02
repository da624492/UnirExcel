package co.gov.sdm.unirExcel.servlet;

import java.io.IOException;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import co.gov.sdm.unirExcel.servicio.UtilExcelAdriana;
import co.gov.sdm.unirExcel.servicio.UtilExcelWilber;

/**
 * Controla la navegacion de la aplicacion.
 * */
@WebServlet(name = "UnirExcelServlet", urlPatterns = {"/servletExcel"})
public class UnirExcelServlet extends HttpServlet {
	@Override
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {		
		String dirOrigen = request.getParameter("directorioEntrada");
		String dirSalida = request.getParameter("directorioSalida");
		
		UtilExcelWilber.extraer(dirOrigen, dirSalida);
		//UtilExcel2021.extraer2021Todos(dirOrigen, dirSalida);
		request.setAttribute("mensaje", "Exceles compilados");
		request.getRequestDispatcher("/paginas/index.jsp").forward(request, response);
	}
}

