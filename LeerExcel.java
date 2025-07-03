/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.salvador.ejercicios.excel_jmm;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author littlejxmmy
 */
public class LeerExcel {

    public static void main(String[] args) throws IOException {

        try (InputStream inp = new FileInputStream("alumnos.xlsx")) {

            //Creamos un nuevo workbook dentro de la hoja de excel "alumnos.xlsx"
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheet("Datos");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            
            //Declaramos las variables
            int suspensos = 0;
            int aprobados = 0;
            int mayor30 = 0;
            
            //Iteramos en las filas de la hoja de excel y la leemos
            for (int i = 1; i < 1001; i++) {

                Row rowDatos = sheet.getRow(i);

                if (rowDatos == null) {

                    continue;

                }
                
                Cell cellDNI = rowDatos.getCell(0);
                Cell cellNombre = rowDatos.getCell(1);
                Cell cellNotas = rowDatos.getCell(2);
                Cell cellEdad = rowDatos.getCell(3);
                System.out.print(cellDNI.getStringCellValue() + " ");
                System.out.print(cellNombre.getStringCellValue() + " ");
                System.out.print(cellNotas.getNumericCellValue() + " ");
                System.out.println(cellEdad.getNumericCellValue() + " ");

                //Creamos un "if if" para hacer dos contadores de notas con los suspensos y aprobados respectivamente
                if (cellNotas.getNumericCellValue() < 5) {

                    suspensos++;

                } if (cellNotas.getNumericCellValue() >= 5) {

                    aprobados++;
                    
                //Comprobamos los alumnos mayores de 30    
                } if (cellEdad.getNumericCellValue() >= 30) {
                    
                    mayor30++;
                
                //Comprobamos que todas las edades se encuentran dentro del rango establecido
                } if (cellEdad.getNumericCellValue() >= 18 && cellEdad.getNumericCellValue() < 60) {
                    
                        
                } else {
                    
                    System.out.println("Hay un valor erróneo en la columna");
                    
                }
            }

            System.out.println();

            //Calculamos la nota media
            Row rowNotaMedia = sheet.createRow(1002);
            Cell cellNotaMedia = rowNotaMedia.createCell(2);
            cellNotaMedia.setCellFormula("AVERAGE(C2:C1001)");
            CellValue cellValueNotaMedia = evaluator.evaluate(cellNotaMedia);
            System.out.println("NOTA MEDIA: " + cellValueNotaMedia.getNumberValue());

            //Calculamos la nota máxima
            Row rowNotaMaxima = sheet.createRow(1003);
            Cell cellNotaMaxima = rowNotaMaxima.createCell(2);
            cellNotaMaxima.setCellFormula("MAX(C2:C1001)");
            CellValue cellValueNotaMaxima = evaluator.evaluate(cellNotaMaxima);
            System.out.println("NOTA MÁXIMA: " + cellValueNotaMaxima.getNumberValue());

            //Calculamos la nota mínima
            Row rowNotaMinima = sheet.createRow(1004);
            Cell cellNotaMinima = rowNotaMinima.createCell(2);
            cellNotaMinima.setCellFormula("MIN(C2:C1001)");
            CellValue cellValueNotaMinima = evaluator.evaluate(cellNotaMinima);
            System.out.println("NOTA MÍNIMA: " + cellValueNotaMinima.getNumberValue());
            
            //Calculamos la edad promedio
            Row rowEdad = sheet.createRow(1005);
            Cell cellEdad = rowEdad.createCell(3);
            cellEdad.setCellFormula("AVERAGE(D2:D1001)");
            CellValue cellValueEdad = evaluator.evaluate(cellEdad);
            System.out.println("EDAD PROMEDIO: " + Math.ceil(cellValueEdad.getNumberValue()) + " años");
            
            //Porcentaje de mayores de 30
            System.out.println("MAYORES DE 30: " + ((mayor30 * 100) / 1000) + "%");

            //Imprimimos por pantalla los aprobados
            System.out.println("APROBADOS: " + aprobados);
            
            //Imprimimos por pantalla los suspensos
            System.out.println("SUSPENSOS: " + suspensos);
            
        }
    }
}
