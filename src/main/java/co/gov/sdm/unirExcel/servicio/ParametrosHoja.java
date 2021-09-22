package co.gov.sdm.unirExcel.servicio;

public class ParametrosHoja {
	private String nombre;	
	private int numFilasEncabezado;
	private int numFilasPieDePagina;
	
	public ParametrosHoja(String nombre, int numFilasEncabezado, int numFilasPieDePagina) {
		this.setNombre(nombre);
		this.setNumFilasEncabezado(numFilasEncabezado);
		this.setNumFilasPieDePagina(numFilasPieDePagina);
	}
	
	public String getNombre() {
		return nombre;
	}	
	private void setNombre(String nombre) {
		this.nombre = nombre;
	}
	
	public int getNumFilasEncabezado() {
		return numFilasEncabezado;
	}
	private void setNumFilasEncabezado(int numFilasEncabezado) {
		this.numFilasEncabezado = numFilasEncabezado;
	}
	
	public int getNumFilasPieDePagina() {
		return numFilasPieDePagina;
	}
	private void setNumFilasPieDePagina(int numFilasPieDePagina) {
		this.numFilasPieDePagina = numFilasPieDePagina;
	}
}
