package com.senai.sp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class ExcelUpdater {

    public void atualizarExcel(){
        Scanner scanner = new Scanner(System.in);

        System.out.print("Digite o caminho do arquivo Excel: ");
        String filePath = scanner.nextLine();
        System.out.print("Digite o nome do arquivo Excel: ");
        String fileName = scanner.nextLine();

        try {
            // Carrega o arquivo existente do Excel
            FileInputStream fileInputStream = new FileInputStream(filePath + "\\" + fileName + ".xlsx");
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Obtém a primeira planilha
            Sheet sheet = workbook.getSheetAt(0);

            // Obtém os dados do usuário
            List<String> rowData = getRowDataFromUser();

            // Adiciona os dados na última linha da planilha
            int lastRowNum = sheet.getLastRowNum();
            Row row = sheet.createRow(lastRowNum + 1);
            int cellNum = 0;
            for (String cellData : rowData) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(cellData);
            }

            // Salva as alterações no arquivo do Excel
            FileOutputStream fileOutputStream = new FileOutputStream(filePath + "\\" + fileName + ".xlsx");
            workbook.write(fileOutputStream);
            workbook.close();
            fileInputStream.close();
            fileOutputStream.close();

            System.out.println("Dados adicionados com sucesso na última linha do arquivo Excel!");

        } catch (IOException e) {
            e.printStackTrace();
        }

        scanner.close();
    }

    private static List<String> getRowDataFromUser() {
        List<String> rowData = new ArrayList<>();

        Scanner scanner = new Scanner(System.in);
        System.out.println("Digite os dados da nova linha:");

        String input;
        while (!(input = scanner.nextLine()).equalsIgnoreCase("fim")) {
            rowData.add(input.trim());
        }

        return rowData;
    }
}
