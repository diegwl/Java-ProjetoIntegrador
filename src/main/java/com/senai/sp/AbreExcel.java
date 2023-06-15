package com.senai.sp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

// Classe que vai realizar todas as ações realizadas no excel e trazer para o Java.
public class AbreExcel {
    private static String fileName = "";

    static List<Produto> listaProdutos = new ArrayList<>();

    // Função que lê o excel,pelo caminho de seu arquivo.
    public static void lerExcel(String arquivoExcel){
        try {
            AbreExcel.fileName = arquivoExcel.replace("\\", "\\\\");
            FileInputStream arquivo = new FileInputStream(new File(AbreExcel.fileName));

            // pega a planilha de produtos.
            XSSFWorkbook workbook = new XSSFWorkbook(arquivo);
            Sheet sheetProdutos = workbook.getSheetAt(0);

            // Olha todas as linhas do planilha
            Iterator<Row> rowIterator = sheetProdutos.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row.getRowNum() == 0){
                    continue;
                } else {

                    // Passa por todas as células da linha atual
                    Iterator<Cell> cellIterator = row.cellIterator();

                    // cria um objeto do Produto presente na linha
                    Produto produto = new Produto();
                    // adiciona o produto na lista de produtos
                    listaProdutos.add(produto);

                    // dependendo do index da coluna, será adicionado o valor no nome do objeto ou no preço
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
            // Tratamento de erros
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Arquivo Excel não encontrado!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    // Função que mostra todos os produtos presentes na lista de produtos
    public static void mostrarProdutos(){
        Iterator<Produto> itr = listaProdutos.iterator();

        System.out.println(" -----------------------------------");
        System.out.printf(" |  %-20s | %-7s |\n", "Nome", "Preço");
        System.out.println(" |-----------------------|---------|");
        while(itr.hasNext())
        {
            Produto produto = itr.next();
            System.out.printf(" |  %-20s | U$ %.2f |\n", produto.getNome(), produto.getPreco());
        }
        System.out.println(" -----------------------------------");
    }

    // Função que mostra o produto que custa mais caro
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
        System.out.printf("\nO maior preço de um produto é:\n%s - U$ %.2f\n", nome, maiorPreco);
    }

    // Função que mostra o produto que custa menos dinheiro
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
        System.out.printf("\nO menor preço de um produto é:\n%s - U$ %.2f\n", nome, menorPreco);
    }

    // Função que mostra a média de preços dos produtos
    public static void mediaPrecos(){
        double soma = 0, media = 0;
        for (Produto produto:listaProdutos) {
            soma = soma + produto.getPreco();
        }
        media = soma / listaProdutos.size();
        System.out.printf("\nA média de preço dos produtos é: U$ %.2f\n", media);
    }

    // Função Main
    public static void main(String[] args) {
        Scanner entrada = new Scanner(System.in);
        System.out.println("Digite o caminho completo do seu arquivo, com o nome já inserido: ");
        String excel = entrada.nextLine();

        // Primero ele chama a função que lê o excel
        lerExcel(excel);

        // Menu para seleção de qual função o usuário deseja ver
        while (true) {
            try {
                System.out.println("\n-----CHAT XLSX PRODUTOS-----\n1 - Lista de produtos\n2 - Maior preço entre os produtos\n3 - Menor preço entre os produtos\n4 - Média de Preços\n5 - Sair\n");
                int opcao = entrada.nextInt();
                switch (opcao) {
                    case 1 -> mostrarProdutos();
                    case 2 -> maiorPreco();
                    case 3 -> menorPreco();
                    case 4 -> mediaPrecos();
                    case 5 -> System.out.println("Finalizando Chat");
                    default -> System.out.println("Opção não existente");
                }
                if (opcao == 5) {
                    entrada.close();
                    break;
                }
            } catch (Exception e) {
                entrada.nextLine();
                System.out.println("Erro " + e);
            }
        }
    }
}
