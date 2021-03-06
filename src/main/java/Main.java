import javax.swing.*;
import java.io.IOException;
import java.util.List;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws IOException {

        GerenciadorCheques gerenciadorCheques = new GerenciadorCheques();

        Scanner sc = new Scanner(System.in);

        int op;

        do {

            System.out.println("1 - Ler e impririr dados do arquivo XLSX");
            System.out.println("2 - Criar e escrever no arquivo XLSX");
            System.out.println("0 - Sair");

            op = sc.nextInt();

            switch (op) {

                case 0:
                    System.exit(0);
                    break;

                case 1:
                    List<Cheque> cheques = gerenciadorCheques.criar();
                    gerenciadorCheques.imprimir(cheques);
                    break;

                case 2:
                    System.out.print("INFORME O NOME PARA O ARQUIVO: ");
                    String nomeArquivo = sc.next();
                    gerenciadorCheques.criaArquivoExcel(nomeArquivo);
                    System.out.println("\nArquivo criado com Sucesso!\n");
                    break;

                default:
                    System.out.println("Opção Incorreta");
            }

        } while (op != 0);

    }
}
