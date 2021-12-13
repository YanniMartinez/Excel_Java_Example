package excel;

//import com.mysql.jdbc.Connection;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import modelo.Conexion;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.dev.XSSFSave;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

    public static void main(String[] args) throws IOException, SQLException {

        crearExcel();
        //leer();
        //cargar();
        //modificar();
        cargarArticulos();
    }

    public static void crearExcel() {
    	System.out.println("Prueba de XLM");
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Prueba");

        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("Hola Mundo");
        row.createCell(1).setCellValue(7.5);
        row.createCell(2).setCellValue(true);

        Cell celda = row.createCell(3);
        celda.setCellFormula(String.format("1+1", ""));

        Row rowUno = sheet.createRow(1);
        rowUno.createCell(0).setCellValue(7);
        rowUno.createCell(1).setCellValue(8);

        Cell celdados = rowUno.createCell(2);
        celdados.setCellFormula(String.format("A%d+B%d", 2, 2));

        try {
            FileOutputStream fileout = new FileOutputStream("Excel.xlsx");
            book.write(fileout);
            fileout.close();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    public static void leer() throws IOException {
        try {
            FileInputStream file = new FileInputStream(new File("D:\\productos.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            int numFilas = sheet.getLastRowNum();

            for (int a = 0; a <= numFilas; a++) {
                Row fila = sheet.getRow(a);
                int numCols = fila.getLastCellNum();

                for (int b = 0; b < numCols; b++) {
                    Cell celda = fila.getCell(b);

                    switch (celda.getCellTypeEnum().toString()) {
                        case "NUMERIC":
                            System.out.print(celda.getNumericCellValue() + " ");
                            break;

                        case "STRING":
                            System.out.print(celda.getStringCellValue() + " ");
                            break;

                        case "FORMULA":
                            System.out.print(celda.getCellFormula() + " ");
                            break;
                    }

                }

                System.out.println("");

            }

        } catch (FileNotFoundException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void cargar() throws IOException, SQLException {

        Conexion con = new Conexion();
        PreparedStatement ps;

        try {
            Connection conn = con.getConexion();
            FileInputStream file = new FileInputStream(new File("C:\\Users\\Yann\\Desktop\\Excel\\Excel\\articulos.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            int numFilas = sheet.getLastRowNum();

            for (int a = 1; a <= numFilas; a++) {
                Row fila = sheet.getRow(a);

                ps = conn.prepareStatement("INSERT INTO producto (codigo, nombre, precio, cantidad) VALUES(?,?,?,?)");
                ps.setString(1, fila.getCell(0).getStringCellValue());
                ps.setString(2, fila.getCell(1).getStringCellValue());
                ps.setDouble(3, fila.getCell(2).getNumericCellValue());
                ps.setDouble(4, fila.getCell(3).getNumericCellValue());
                ps.execute();
            }
            
            conn.close();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void modificar() throws IOException {
        try {
            FileInputStream file = new FileInputStream(new File("D:\\productos.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            XSSFRow fila = sheet.getRow(1); //Obtiene la fila 1
            
            if(fila == null) //Si no existe la fila, la crea
            {
               fila = sheet.createRow(1);
            }
            
            XSSFCell celda = fila.createCell(1); //Obtiene la celda 1
            
            if(celda == null) //Si no existe la celda la crea
            {
               celda = fila.createCell(1);
            }
            
            celda.setCellValue("Modificacion"); //ASigna el nuevo valor
            
            file.close(); //Cierra la hoja
            
            FileOutputStream output = new FileOutputStream("D:\\nuevo.xlsx"); //Puede ser en el mismo archivo o en otro 
            wb.write(output); //En esta parte se envia todo lo que se hizo arriba
            output.close(); //Cierra
            

        } catch (FileNotFoundException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    
    public static void cargarArticulos() throws IOException, SQLException {
    	
    	//Primera comilla va ruta y nombre de la base de datos a la que nos queremos conectar
    	//Segunda Nombre de usuario
    	//Tercera es la contraseña de la base de datos
    	Connection cn = DriverManager.getConnection("jdbc:mysql://localhost/xcommerce","root","admin");
        /*PreparedStatement pst = cn.prepareStatement("Insert into alumnos values (?,?,?)"); //Aquí va la consulta
    	pst.setString(1, "0"); //el primer campo hace referencia a la columna de la BD y el segundo al valor
    	pst.setString(2, "Un nombre de alumno"); //Para la columna del alumno   Metodo: trim() elimina espacios al inicio y al final.
    	pst.setString(3, "Un grupo del alumno"); //el
    	pst.executeUpdate(); //Ejecutará la consulta*/
    	
    	PreparedStatement pst = cn.prepareStatement("select * from articles where id=?"); //Selecciona todo del id que recibirá
    	//Enviandole a la base de datos lo que queremos consultar.
    	pst.setString(1, "1"); //Se le indica el numero del campo y el valor a buscar.
    	
    	//Nos ayudará a obtener el resultado de la consulta.
    	ResultSet rs = pst.executeQuery(); //Nos permitirá saber si se encontraron o no los resultados.
    	//validando si obtuvo valores o no:
    	if(rs.next()) { //Si encontró algo entonces:
    		System.out.println( rs.getString("category") ); //Va entre comilla el campo que queremos ver
    		System.out.println( rs.getString("description")); //Obtiene la descripción
    	}else {
    		System.out.println("No se obtuvieron registros");
    	}
    	
    }
    
}
