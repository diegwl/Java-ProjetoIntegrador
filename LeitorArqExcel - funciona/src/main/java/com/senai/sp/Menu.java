package com.senai.sp;

import java.util.Scanner;

public class Menu {
    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);
        String answer = "S";
        int opcao = 0;

        while(true){
            System.out.println("Menu");
            System.out.println("1 - Abrir");
            System.out.println("2 - Criar");
            System.out.println("3 - Alterar");
            System.out.println("4 - Sair");
            System.out.println("Digite a opção desejada:");
            opcao = sc.nextInt();
            switch(opcao){
                case 1:
                    ExcelReader er = new ExcelReader();
                    er.lerExcel();
                    break;
                case 2:
                    ExcelWriter ew = new ExcelWriter();
                    ew.criarExcel();
                    break;
                case 3:
                    ExcelUpdater eu = new ExcelUpdater();
                    eu.atualizarExcel();
                    break;
                case 4:
                    System.out.println("Deseja finalizar o programa? S ou N");
                    answer = sc.next().toLowerCase();
                    break;
                default:
                    System.out.println("Opção inválida!");
                    break;
            }
            if(answer.equals("s")){
                break;
            }
        }

        /*Scanner scanner = new Scanner(System.in);
        boolean sair = false;

        while (true) {
            exibirMenu();
            System.out.println("Digite a opção desejada: ");
            int opcao = scanner.nextInt();

                switch (opcao) {
                    case 1:
                        System.out.println("Opção escolhida: Criar");
                        ExcelWriter ew = new ExcelWriter();
                        ew.criarExcel();
                        System.out.println("Pressione Enter para continuar...");
                        scanner.nextLine();
                        break;
                    case 2:
                        System.out.println("Opção escolhida: Abrir");
                        ExcelReader er = new ExcelReader();
                        er.lerExcel();
                        System.out.println("Pressione Enter para continuar...");
                        scanner.nextLine();
                        break;
                    case 3:
                        System.out.println("Opção escolhida: Alterar");
                        ExcelUpdater eu = new ExcelUpdater();
                        eu.atualizarExcel();
                        System.out.println("Pressione Enter para continuar...");
                        scanner.nextLine();
                        break;
                    case 4:
                        System.out.println("Opção escolhida: Sair");
                        sair = true;
                        break;
                    default:
                        System.out.println("Opção inválida. Digite novamente.");
                        break;
                }
           // }
        }

        //System.out.println("Programa encerrado.");*/
    }

    public static void exibirMenu() {
        System.out.println("===== MENU =====");
        System.out.println("1 - Criar");
        System.out.println("2 - Abrir");
        System.out.println("3 - Alterar");
        System.out.println("4 - Sair");
    }
}
