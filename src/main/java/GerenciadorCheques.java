import lombok.Cleanup;
import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;

public class GerenciadorCheques {

    private Scanner sc = new Scanner(System.in);

    public List<Cheque> criar() throws IOException {

        List<Cheque> cheques = new ArrayList<>();

        @Cleanup FileInputStream file = new FileInputStream("src/main/resources/controle_cheques.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        List<Row> rows = (List<Row>) toList(sheet.iterator());

        rows.remove(0);

        rows.forEach(row -> {

            List<Cell> cells = (List<Cell>) toList(row.cellIterator());

            Cheque cheque = Cheque.builder()

                    .data(cells.get(0).getDateCellValue())

                    .numeroCheque((int) cells.get(1).getNumericCellValue())

                    .nome(cells.get(2).getStringCellValue())

                    .valor(new BigDecimal(cells.get(3).getNumericCellValue()))

                    .status(cells.get(4).getStringCellValue())

                    .qtParcelas((int) cells.get(5).getNumericCellValue())

                    .formulaGeral(cells.get(6).getCellFormula())

                    .build();

            cheques.add(cheque);

        });

        return cheques;
    }


    public void criaArquivoExcel(String nomeArquivo) {

        Workbook workbook = new XSSFWorkbook();

        criaPlanilha(workbook);

        try (OutputStream fileOut = new FileOutputStream("src/main/resources/" + nomeArquivo + ".xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public void criaPlanilha(Workbook workbook) {

        System.out.print("Informe o nome da Planilha: ");

        String nomePlanilha = sc.next();

        Sheet sheet = workbook.createSheet(nomePlanilha);

        criaLinha(sheet);

    }

    public void criaLinha(Sheet sheet) {

        Row row = sheet.createRow(0);

        System.out.println("Para realização de Testes foi criado para UMA LINHA na planilha = " + sheet.getSheetName());

        criaCelulas(row);

    }

    public void criaCelulas(Row row) {

        System.out.print("Informe a quantidade e Células:");

        int quantidadeCelulas = sc.nextInt();

        for (int posicao = 0; posicao < quantidadeCelulas; posicao++) {
            insereUmaCelula(row, posicao);
        }
    }

    public void insereUmaCelula(Row row, Integer posicao) {

        System.out.print("Informe o texto da celula na posicao = " + (posicao + 1) + " : ");
        Cell cell = row.createCell(posicao);
        String textoCelula = sc.next();
        cell.setCellValue(textoCelula);

    }

    public List<?> toList(Iterator<?> iterator) {
        return IteratorUtils.toList(iterator);
    }

    public void imprimir(List<Cheque> cheques) {
        cheques.forEach(System.out::println);
    }


}
