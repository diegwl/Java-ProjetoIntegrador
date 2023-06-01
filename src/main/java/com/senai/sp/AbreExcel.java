package com.senai.sp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.*;

public class AbreExcel {
    private static final String fileName = "C:\\Users\\46874383850\\Desktop\\Produtos.xlsx";

    static List<Produto> listaProdutos = new ArrayList<>();
    public static void lerExcel(){
        try {
            FileInputStream arquivo = new FileInputStream(new File(AbreExcel.fileName));

            XSSFWorkbook workbook = new XSSFWorkbook(arquivo);
            Sheet sheetProdutos = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheetProdutos.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row.getRowNum() == 0){
                    continue;
                } else {
                    Iterator<Cell> cellIterator = row.cellIterator();

                    Produto produto = new Produto();
                    listaProdutos.add(produto);
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getColumnIndex()) {
                            case 0 -> produto.setNome(cell.getStringCellValue());
                            case 1 -> {
                                String precoStr = cell.getStringCellValue();
                                String precoStr2 = precoStr.replace("U$ ", "").replace(",", ".");
                                double preco = Double.parseDouble(precoStr2);
                                produto.setPreco(preco);
                            }
                        }
                    }
                }
            arquivo.close();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Arquivo Excel não encontrado!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void mostrarProdutos(){
        Iterator<Produto> itr = listaProdutos.iterator();

        while(itr.hasNext())
        {
            Produto produto = itr.next();
            System.out.printf("%s - U$ %.2f\n", produto.getNome(), produto.getPreco());
        }
    }

    public static void maiorPreco(){
        double maiorPreco = -1;
        String nome = null;
        for (Produto produto:listaProdutos) {
            if (maiorPreco == -1){
                maiorPreco = produto.getPreco();
                nome = produto.getNome();
            } else if (produto.getPreco() > maiorPreco){
                maiorPreco = produto.getPreco();
                nome = produto.getNome();
            }
        }
        System.out.printf("\nO maior preço de um produto é:\n%s - U$ %.2f", nome, maiorPreco);
    }

    public static void menorPreco(){
        double menorPreco = -1;
        String nome = null;
        for (Produto produto:listaProdutos) {
            if (menorPreco == -1) {
                menorPreco = produto.getPreco();
                nome = produto.getNome();
            } else if (produto.getPreco() < menorPreco) {
                menorPreco = produto.getPreco();
                nome = produto.getNome();
            }
        }
        System.out.printf("\nO menor preço de um produto é:\n%s - U$ %.2f", nome, menorPreco);
    }

    public static void mediaPrecos(){
        double soma = 0, media = 0;
        for (Produto produto:listaProdutos) {
            soma = soma + produto.getPreco();
        }
        media = soma / listaProdutos.size();
        System.out.printf("\nA média de preço dos produtos é: U$ %.2f", media);
    }

    public static void main(String[] args) {
        lerExcel();
        maiorPreco();
        menorPreco();
        mediaPrecos();
    }
}
