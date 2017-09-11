package es.mercadon.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * http://poi.apache.org/download.html#POI-3.15
 * https://www.tutorialspoint.com/apache_poi/apache_poi_core_classes.htm
 * https://www.tutorialspoint.com/apache_poi/apache_poi_spreadsheets.htm
 * @author llizamab
 *
 */
public class Prueba {

	// INCONF-PLP-ED_InventarioDispositivo_v1_0_4.xls
	// OPVTA-PLP-ED_TramosParkingDefecto_v1_0_0.xls

	// codigo de la peticion
//	public static final String COD_PET = "MAQ-XX55";
	// nombre del PLP
//	public static final String PLP = "INCONF-PLP-ED_InventarioDispositivo_v1_0_4.xls";
//	public static final String PLP = "OPVTA-PLP-ED_TramosParkingDefecto_v1_0_0.xls";
	
//	public static final String RUTA = "S:\\Workspace\\eclipse\\fwk30\\prueba\\excel\\";
//	public static final String RUTA_XML = "S:\\Workspace\\eclipse\\fwk30\\prueba\\xml\\";
	
	// pendiente implementar login
	final static Logger logger = Logger.getLogger(Prueba.class);
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// leer properties
		final File properties = new File("S:\\TestLinkGenerator3000\\config.properties");
		// si no existe chao
		if (!properties.isFile() || !properties.exists()) {
			logger.error("Fichero de configuracion no existe!");
			return;
		}
		// leo variables
		
//		final File file = new File(RUTA + PLP);

		try {
			final Properties prop = new Properties();
			           
			final InputStream stream = new FileInputStream(properties);
			prop.load(stream);
			
			final String COD_PET = prop.getProperty("COD_PET");
			final String PLP = prop.getProperty("PLP");
			final String RUTA_XML = prop.getProperty("RUTA_XML");
//			# columna O 14 - Base salida
			int intBaseSalida = 14;
//			# columna J 9 - Base de entrada
			int intBaseEntrada = 9;
//			# columna H 7 - BBDD (precondicion)
			int intBbdd = 7;
//			# columna G 6 - Descripción
			int intDesc = 6;
//			# columna C 2 - TC ID
			int intTcId = 2;
//			# columna B 1 - REQUERIMIENTO
			int intReqFun = 1;
			// cargo los valores
			final String BASE_SALIDA = prop.getProperty("BASE_SALIDA");
			// si no es null ni vacio
			if (BASE_SALIDA != null && !BASE_SALIDA.trim().isEmpty()) {
				intBaseSalida = COLUMNAS.get(BASE_SALIDA);
			}
			final String BASE_ENTRADA = prop.getProperty("BASE_ENTRADA");
			// si no es null ni vacio
			if (BASE_ENTRADA != null && !BASE_ENTRADA.trim().isEmpty()) {
				intBaseEntrada = COLUMNAS.get(BASE_ENTRADA);
			}
			final String BB_DD = prop.getProperty("BB_DD");
			// si no es null ni vacio
			if (BB_DD != null && !BB_DD.trim().isEmpty()) {
				intBbdd = COLUMNAS.get(BB_DD);
			}
			final String DESCRIPCION = prop.getProperty("DESCRIPCION");
			// si no es null ni vacio
			if (DESCRIPCION != null && !DESCRIPCION.trim().isEmpty()) {
				intDesc = COLUMNAS.get(DESCRIPCION);
			}
			final String TC_ID = prop.getProperty("TC_ID");
			// si no es null ni vacio
			if (TC_ID != null && !TC_ID.trim().isEmpty()) {
				intTcId = COLUMNAS.get(TC_ID);
			}
			final String REQUERIMIENTO = prop.getProperty("REQUERIMIENTO");
			// si no es null ni vacio
			if (REQUERIMIENTO != null && !REQUERIMIENTO.trim().isEmpty()) {
				intReqFun = COLUMNAS.get(REQUERIMIENTO);
			}
			
			logger.error("Variables de configuracion: " + prop);

			final File file = new File(PLP);

			logger.error("Fichero a procesar: " + file.getName());
			
			if (file.isFile() && file.exists()) {

				final InputStream fIP = new FileInputStream(file);
				// Get the workbook instance for XLSX file
				// XSSFWorkbook workbook = new XSSFWorkbook(fIP);
				final HSSFWorkbook workbook = new HSSFWorkbook(fIP);
				
				// creo el Testsuite de la peticion
				final Testsuite suitePet = new Testsuite(COD_PET, "Plan de pruebas de la peticion " + COD_PET);
				
				logger.error("Generando plan de pruebas de la peticion " + COD_PET);

				final Iterator<Sheet> iterator = workbook.iterator();
				// por cada
				while (iterator.hasNext()) {
					final Sheet sheet = iterator.next();

					// si es hoja de operacion
					if (sheet.getSheetName().matches("^OP_[0-9]{1,2}$")) {
						
//						logger.error("Hoja:" + sheet.getSheetName());
						
						final Cell cell0 = sheet.getRow(0).getCell(1);
						
						String nombreOperacion = (cell0 != null) ? cell0.getStringCellValue() : null;
						// si es null busco en la linea siguiente
						if (nombreOperacion == null || nombreOperacion.isEmpty()) {
							final Cell cell1 = sheet.getRow(1).getCell(1);
							nombreOperacion = (cell1 != null) ? cell1.getStringCellValue() : null;
							// si sigue siendo null, informo error de formato y paso a la siguiente operacion
							if (nombreOperacion == null) {
								
								logger.error("No se ha encontrado el nombre de la operacion al procesar la hoja: " + sheet.getSheetName());
								logger.error("Revisar el formato del fichero");
								continue;
							}
						}

						nombreOperacion = nombreOperacion.replace("Operación ", "")
								.replace("Operacion ", "");
						
						// creo la suite de la operacion
						
						final Testsuite suiteOp = new Testsuite(nombreOperacion, "Plan de pruebas de la operacion " + nombreOperacion);
						
						suitePet.addTestSuite(suiteOp);
						
						logger.error("Generando casos de la operacion: " + nombreOperacion);
						// buscar nombre de operacion
						final Iterator<Row> rowIterator = sheet.iterator();
						final Iterator<Row> rowIterator2 = sheet.iterator();
						int cont = 0;
						
						// me muevo a la final 9 <--------- cambiar esto para que se mueva hasta pillar un testcase (algunos plps tienen leyenda)
						while (rowIterator2.hasNext()) {
							final Row row =  rowIterator2.next();
							// si la final tiene info es un REQ-FUN dejo de avanzar
							if (row.getCell(1) != null && row.getCell(1).getStringCellValue() != null 
								&& !row.getCell(1).getStringCellValue().trim().isEmpty()) {
								// valor
								final String valor = row.getCell(1).getStringCellValue().trim();
								// si encuentro requ fun
								if (valor.matches("^REQ-FUN-.+$")) {
									// dejo de avanzar
									break;
								}
							}
							cont++;
						}
						// me muevo lo que corresponda
						for (int x = 0; x < cont; x ++) {
							rowIterator.next();
						}
						// comienzo a leer los casos
						while (rowIterator.hasNext()) {
							final Row row =  rowIterator.next();
							// si la final tiene info
							if (row.getCell(1) != null && row.getCell(1).getStringCellValue() != null 
									&& !row.getCell(1).getStringCellValue().trim().isEmpty()) {
								// columna B 1 - REQUERIMIENTO
//								logger.error("REQUERIMIENTO: " + row.getCell(1).getStringCellValue());
								final String reqFun = row.getCell(intReqFun).getStringCellValue();
								// columna C 2 - TC ID
								logger.error("TC ID: " + row.getCell(intTcId).getStringCellValue());
								final String tcId = row.getCell(intTcId).getStringCellValue();
								// columna G 6 - Descripción
//								logger.error("Descripción: <" + row.getCell(6).getStringCellValue() + ">");
								final String descripcion = row.getCell(intDesc).getStringCellValue();
								// columna H 7 - BBDD (precondicion)
//								logger.error("BBDD: " + row.getCell(7).getStringCellValue());
								final String precond = row.getCell(intBbdd).getStringCellValue();
								// columna J 9 - Base de entrada
//								logger.error("Base de entrada: " + row.getCell(9).getStringCellValue());
								final String baseIn = row.getCell(intBaseEntrada).getStringCellValue();
								// columna O 14 - Base salida
//								logger.error("Base salida: " + row.getCell(14).getStringCellValue());
								final String baseOut = row.getCell(intBaseSalida).getStringCellValue();
								
								final TestCase testCase = new TestCase();
								testCase.name = tcId + " - " + nombreOperacion;
								testCase.preconditions = precond + "\nParámetros de entrada: " + baseIn;
								// si la descripcion contiene mas de una linea
								if (descripcion.contains("\n")) {
									testCase.summary = reqFun + ": " + descripcion.substring(0, descripcion.indexOf("\n"));
								} else {
									testCase.summary = reqFun + ": " + descripcion;
								}
								// agrego cada testCase
								suiteOp.addTestCase(testCase);
								// agrego paso
								final Step step = new Step();
								// si contiene salto de linea
								if (descripcion.contains("\n")) {
									step.actions = descripcion.substring(descripcion.indexOf("\n"));
								} else {
									step.actions = descripcion;
								}
								step.expectedresults = baseOut;
								testCase.addStep(step);
							}

//							Iterator<Cell> cellIterator = row.cellIterator();
//							while (cellIterator.hasNext()) {
//								Cell cell = cellIterator.next();
////								logger.error(cell.getStringCellValue());
//							}
						};
					}
				}
				// generar el xml
				final String xml = suitePet.generarXml();
				final String nombreXml = "TL_Suite_" + suitePet.name + ".xml";
				
				final FileWriter fooWriter = new FileWriter(RUTA_XML + nombreXml, false);
				fooWriter.append(xml);
				fooWriter.close();
				// generar fichero
				logger.error("Fichero " + nombreXml + " generado correctamente en la ruta: " + RUTA_XML);

				workbook.close();
			} else {
				// la ruta no existe
				logger.error("No existe al ruta o el fichero PLP");
			}

		} catch (final FileNotFoundException e) {
			// TODO Auto-generated catch block
			logger.error("FileNotFoundException : "+ e.getMessage(), e);
			
		} catch (final IOException e) {
			// TODO Auto-generated catch block
			logger.error("IOException : "+ e.getMessage(), e);
		} catch (final Exception e) {
			logger.error("Exception : "+ e.getMessage(), e);
		}
	}
	
	public static class Testsuite {
		public String name = null;
		public String details = null;
		public List<Testsuite> tests = new ArrayList<Testsuite>();
		public List<TestCase> testCases = new ArrayList<TestCase>();
		public StringBuilder sbd = new StringBuilder();
		
		Testsuite(String nombre, String details) {
			this.name = nombre;
			this.details = details;
		}
		
		public void addTestSuite(Testsuite test) {
			this.tests.add(test);
		}
		
		public void addTestCase(TestCase testCase) {
			this.testCases.add(testCase);
		}
		
		public String generarXml() {
			// salto de linea
			final String newLine = "\n";
			// <?xml version="1.0" encoding="UTF-8"?>
			sbd.append("<?xml version=\"1.0\" encoding=\"iso-8859-1\"?>" + newLine);
			// <testsuite name="">
			sbd.append("<testsuite name=\"" + this.name + "\">" + newLine);
		    // <details><![CDATA[]]></details>
			sbd.append("<details><![CDATA[")
			.append(this.details)
			.append("]]></details>" + newLine);
			// por cada testSuite
			for (final Testsuite test : this.tests) {
				// <testsuite name="">
				sbd.append("\t<testsuite name=\"" + test.name + "\">" + newLine);
			    // <details><![CDATA[]]></details>
				sbd.append("\t<details><![CDATA[")
				.append(test.details)
				.append("]]></details>" + newLine);
				// por cada testCase
				for (final TestCase testCase : test.testCases) {
					// <testcase name="">
					sbd.append("\t\t<testcase name=\"" + testCase.name + "\">" + newLine);
					// <summary><![CDATA[]]></summary>
					sbd.append("\t\t\t<summary><![CDATA[" + testCase.summary + "]]></summary>" + newLine);
					// <preconditions><![CDATA[]]></preconditions>
					sbd.append("\t\t\t<preconditions><![CDATA[<p>" + testCase.preconditions.replace("\n", "</br>") + "</p>]]></preconditions>" + newLine);
					// <execution_type><![CDATA[1]]></execution_type>
					sbd.append("\t\t\t<execution_type><![CDATA[1]]></execution_type>" + newLine);
					// <steps>
					sbd.append("\t\t\t<steps>" + newLine);
					// por cada step
					for (final Step step : testCase.steps) {
						// <step>
						sbd.append("\t\t\t\t<step>" + newLine);
						// <step_number><![CDATA[1]]></step_number>
						sbd.append("\t\t\t\t<step_number><![CDATA[" + step.step_number + "]]></step_number>" + newLine);
						// <actions><![CDATA[]]></actions>
						sbd.append("\t\t\t\t<actions><![CDATA[<p>" + step.actions.replace("\n", "</br>") + "</p>]]></actions>" + newLine);
						// <expectedresults><![CDATA[]]></expectedresults>
						sbd.append("\t\t\t\t<expectedresults><![CDATA[" + step.expectedresults + "]]></expectedresults>" + newLine);
						// <execution_type><![CDATA[1]]></execution_type>
						sbd.append("\t\t\t\t<execution_type><![CDATA[1]]></execution_type>" + newLine);
						// </step>
						sbd.append("\t\t\t\t</step>" + newLine);
					}
					// </steps>
					sbd.append("\t\t\t</steps>" + newLine);
					// </testcase>
					sbd.append("\t\t</testcase>" + newLine);
				}
				// </testsuite>
				sbd.append("\t</testsuite>" + newLine);
			}
			// </testsuite>
			sbd.append("</testsuite>");
			// retorno
			return sbd.toString();
		}
	}

	public static class TestCase {
		
		public String name = null;
		public String summary = null;
		public String preconditions = null;
		
		public List<Step> steps = new ArrayList<Step>();
		
		public void addStep(Step step) {
			this.steps.add(step);
		}
	}
	
	public static class Step {
		
		public int step_number = 1;
		public String actions = null;
		public String expectedresults = null;
		
	}
	
	public static Map<String, Integer> COLUMNAS = new HashMap<String, Integer>() {
		private static final long serialVersionUID = 1L;

	{
		put("A", 0);
		put("B", 1);
		put("C", 2);
		put("D", 3);
		put("E", 4);
		put("F", 5);
		put("G", 6);
		put("H", 7);
		put("I", 8);
		put("J", 9);
		put("K", 10);
		put("L", 11);
		put("M", 12);
		put("N", 13);
		put("O", 14);
		put("P", 15);
		put("Q", 16);
		put("R", 17);
		put("S", 18);
		put("T", 19);
		put("U", 20);
		put("V", 21);
		put("W", 22);
		put("X", 23);
		put("Y", 24);
		put("Z", 25);
	}};
}
