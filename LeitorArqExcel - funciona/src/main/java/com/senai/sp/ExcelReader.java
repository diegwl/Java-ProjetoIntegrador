package com.senai.sp;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

        public void lerExcel(){
            Scanner scanner = new Scanner(System.in);

            System.out.print("Digite o caminho do arquivo Excel: ");
            String filePath = scanner.nextLine();
            System.out.println("Enter the filename:");
            String filename = scanner.nextLine();

            try {
                FileInputStream file = new FileInputStream(filePath+"\\"+filename);

                // Carrega o arquivo do Excel usando o Apache POI
                Workbook workbook = new XSSFWorkbook(file);

                // Obtém a primeira planilha
                Sheet sheet = workbook.getSheetAt(0);

                // Lista para armazenar as linhas da planilha
                List<List<String>> tableData = new ArrayList<>();

                // Itera sobre as linhas da planilha
                for (Row row : sheet) {
                    List<String> rowData = new ArrayList<>();

                    // Itera sobre as células de cada linha
                    for (Cell cell : row) {
                        // Obtém o valor da célula
                        String cellValue = formatCellValue(cell);

                        // Adiciona o valor da célula à lista de dados da linha
                        rowData.add(cellValue);
                    }

                    // Adiciona a lista de dados da linha à tabela
                    tableData.add(rowData);
                }

                // Fecha o arquivo
                workbook.close();
                file.close();

                // Exibe a tabela no console
                displayTable(tableData);
            } catch (IOException e) {
                e.printStackTrace();
            }

            scanner.close();
        }

        private static String formatCellValue(Cell cell) {
            String cellValue = "";
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    cellValue = cell.getCellFormula();
                    break;
                default:
                    break;
            }
            return cellValue;
        }

        private static void displayTable(List<List<String>> tableData) {
            // Verifica o tamanho máximo de cada coluna
            int[] columnWidths = new int[tableData.get(0).size()];
            for (List<String> rowData : tableData) {
                for (int i = 0; i < rowData.size(); i++) {
                    if (rowData.get(i).length() > columnWidths[i]) {
                        columnWidths[i] = rowData.get(i).length();
                    }
                }
            }

            // Imprime a tabela formatada
            for (List<String> rowData : tableData) {
                for (int i = 0; i < rowData.size(); i++) {
                    String cellValue = rowData.get(i);
                    int padding = columnWidths[i] - cellValue.length();
                    System.out.print(cellValue);
                    for (int j = 0; j < padding; j++) {
                        System.out.print(" ");
                    }
                    System.out.print("   "); // Espaçamento entre colunas
                }
                System.out.println(); // Quebra de linha após cada linha da tabela
            }
        }
}
