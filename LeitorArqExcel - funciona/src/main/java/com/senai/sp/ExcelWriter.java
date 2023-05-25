package com.senai.sp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class ExcelWriter {

    public void criarExcel(){
        Scanner scanner = new Scanner(System.in);

        System.out.print("Digite o caminho do arquivo Excel: ");
        String filePath = scanner.nextLine();
        System.out.print("Digite o nome do arquivo Excel: ");
        String fileName = scanner.nextLine();

        try {
            // Cria um novo arquivo do Excel
            Workbook workbook = new XSSFWorkbook();

            // Cria uma planilha no arquivo
            Sheet sheet = workbook.createSheet("Dados");

            // Obtém os dados do usuário
            List<List<String>> tableData = getTableDataFromUser();

            // Escreve os dados na planilha
            writeDataToSheet(sheet, tableData);

            // Salva o arquivo do Excel
            saveExcelFile(workbook, filePath, fileName);

            System.out.println("Dados inseridos com sucesso no arquivo Excel!");

        } catch (IOException e) {
            e.printStackTrace();
        }

        scanner.close();
    }

    private static List<List<String>> getTableDataFromUser() {
        List<List<String>> tableData = new ArrayList<>();

        Scanner scanner = new Scanner(System.in);
        System.out.println("Digite os dados da tabela (Digite 'fim' para finalizar a entrada de dados):");

        String input;
        while (!(input = scanner.nextLine()).equalsIgnoreCase("fim")) {
            String[] values = input.split(",");
            List<String> rowData = new ArrayList<>();
            for (String value : values) {
                rowData.add(value.trim());
            }
            tableData.add(rowData);
        }

        return tableData;
    }

    private static void writeDataToSheet(Sheet sheet, List<List<String>> tableData) {
        int rowNum = 0;
        for (List<String> rowData : tableData) {
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;
            for (String cellData : rowData) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(cellData);
            }
        }
    }

    private static void saveExcelFile(Workbook workbook, String filePath, String fileName) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(filePath + "\\" + fileName + ".xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }
}
