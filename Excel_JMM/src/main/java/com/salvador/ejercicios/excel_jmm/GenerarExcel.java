package com.salvador.ejercicios.excel_jmm;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import java.util.Random;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 *
 * @author littlejxmmy
 */
public class GenerarExcel {

    /**
     * Declaramos el método main donde llamamos al método generarExcel()
     *
     * @param args
     */
    public static void main(String[] args) {

        generarExcel();
        
    }

    /**
     * Declaramos el método generarExcel donde creamos el excel
     */
    public static void generarExcel() {
        
        //Creamos la hoja "Datos"
        Workbook wb = new XSSFWorkbook();//
        Sheet sheet = wb.createSheet("Datos");
        Sheet sheetEstadisticas = wb.createSheet("Estadísticas");
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Creamos la fila principal de la hoja "Datos"
        Row headerrow = sheet.createRow(0);
        headerrow.createCell(0).setCellValue("DNI");
        headerrow.createCell(1).setCellValue("Nombre");
        headerrow.createCell(2).setCellValue("Nota");
        headerrow.createCell(3).setCellValue("Edad");
        
        //Creamos las filas de la hoja "Estadísticas"
        Row rowEstadisticas = sheetEstadisticas.createRow(0);
        rowEstadisticas.createCell(0).setCellValue("Estadísticas");
        
        Row rowAlumnos = sheetEstadisticas.createRow(1);
        rowAlumnos.createCell(0).setCellValue("Total Alumnos");
        
        Row rowAprobados = sheetEstadisticas.createRow(2);
        rowAprobados.createCell(0).setCellValue("Aprobados"); 
        
        Row rowSuspensos = sheetEstadisticas.createRow(3);
        rowSuspensos.createCell(0).setCellValue("Suspensos");
        
        Row rowPromedio = sheetEstadisticas.createRow(4);
        rowPromedio.createCell(0).setCellValue("Promedio Notas");

        //Array con DNI aleatorios y letras
        Random randomDNI = new Random();
        String[] letras = {"T", "R", "W", "A", "G", "M", "Y", "F", "P", "D", "X", "B", "N", "J", "Z", "S", "Q", "V", "H", "L", "C", "K", "E"};
        int[] DNI = new int[1001];

        //Array predefinido con los nombres de los alumnos
        String[] alumnos = {"Elena", "Jose", "Margarita", "Alberto", "Marcos", "Antonio", "María", "Ainhoa", "Luís", "Nerea"};
        Random randomAlumnos = new Random();

        //Array con notas aleatorias
        Random randomNotas = new Random();
        double[] notas = new double[1001];
        
        //Creamos el array con las edades
        Random randomEdad = new Random();
        int[] edad = new int[1001];
        

        for (int i = 1; i < 1001; i++) {

            Row rowDatos = sheet.createRow(i);
            DNI[i] = randomDNI.nextInt(10000000, 99999999);
            notas[i] = randomNotas.nextDouble(0, 10);
            edad[i] = (int) randomEdad.nextDouble(18, 60);

            //Creamos las filas con todos los DNI y sus letras
            Cell cellDNI = rowDatos.createCell(0);
            int letra = DNI[i] % 23;
            cellDNI.setCellValue(DNI[i] + letras[letra]);
            System.out.print(DNI[i] + letras[letra] + " ");

            //Creamos las filas con sus nombres aleatorios
            Cell cellNombre = rowDatos.createCell(1);
            cellNombre.setCellValue(alumnos[randomAlumnos.nextInt(alumnos.length)]);
            System.out.print(alumnos[randomAlumnos.nextInt(alumnos.length)] + " ");

            //Creamos las filas con las notas aleatorias
            Cell cellNotas = rowDatos.createCell(2);
            cellNotas.setCellValue(Math.ceil(notas[i]));
            System.out.print(Math.ceil(notas[i]));
            
            //Creamos las filas de la columna edad
            Cell cellEdad = rowDatos.createCell(3);
            cellEdad.setCellValue(edad[i]);
            System.out.println(" " + edad[i]);

        }
        
        try {

            FileOutputStream fileout = new FileOutputStream("alumnos.xlsx");
            wb.write(fileout);
            fileout.close();
            File fileExist = new File("alumnos.xlsx");
            
            System.out.println();
            
            if (fileExist.exists() && fileExist.canWrite()) {
                    
                System.out.println("El archivo alumnos.xlsx existe en el sistema y se puede sobreescribir");
                
            } else {

                System.out.println("Error no se ha podido realizar la petición");

            }
            

        } catch (FileNotFoundException ex) {

            Logger.getLogger(GenerarExcel.class.getName()).log(Level.SEVERE, null, ex);

        } catch (IOException ex) {

            Logger.getLogger(GenerarExcel.class.getName()).log(Level.SEVERE, null, ex);

        }
    }
    
    /*public static int DNI() {
        
        int DNI;
        Random randomDNI = new Random();
        DNI = randomDNI.nextInt(10000000, 99999999);
        
        return DNI;
    }
    
    public static String letraDNI(int DNI) {
        
        int letra;
        String letraValida;
        String[] letras = {"T", "R", "W", "A", "G", "M", "Y", "F", "P", "D", "X", "B", "N", "J", "Z", "S", "Q", "V", "H", "L", "C", "K", "E"};
            
        letra = DNI % letras.length;
        letraValida = letras[letra];
        
        return letraValida;
    }
    
    public static String alumnos() {
        
        String alumnos;
        String[] nombres = {"Elena", "Jose", "Margarita", "Alberto", "Marcos", "Antonio", "María", "Ainhoa", "Luís", "Nerea"};
        Random randomAlumnos = new Random();
        alumnos = nombres[randomAlumnos.nextInt(nombres.length)];
        
        return alumnos;
    }*/
}
