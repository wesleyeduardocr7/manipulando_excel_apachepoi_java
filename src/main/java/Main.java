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

            System.out.println("1 - Ler e impririr dados do arquivo xlsx");
            System.out.println("2 - Criar e escrever no arquivo xlsx");
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
                    System.out.println("Informe o nome para o  arquivo");
                    String nomeArquivo = sc.next();
                    gerenciadorCheques.criaArquivoExcel(nomeArquivo);
                    break;

                default:
                    System.out.println("Opção Incorreta");
            }

        } while (op != 0);

    }
}
