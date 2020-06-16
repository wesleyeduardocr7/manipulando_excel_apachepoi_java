import lombok.Cleanup;
import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class GerenciadorCheques {

    public List<Cheque> criar() throws IOException {

        List<Cheque> cheques = new ArrayList<>();

        @Cleanup FileInputStream file = new FileInputStream("src/main/resources/controle_cheques.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet =  workbook.getSheetAt(0);

        List<Row> rows = (List<Row>) toList(sheet.iterator());

        rows.remove(0);

        rows.forEach(row ->{

            List<Cell> cells = (List<Cell>) toList(row.cellIterator());

            Cheque cheque = Cheque.builder()

                    .data(cells.get(0).getDateCellValue())

                    .numeroCheque( (int)cells.get(1).getNumericCellValue())

                    .nome(cells.get(2).getStringCellValue())

                    .valor(new BigDecimal(cells.get(3).getNumericCellValue()))

                    .status(cells.get(4).getStringCellValue())

                    .qtParcelas( (int)cells.get(5).getNumericCellValue())

                    .formulaGeral(cells.get(6).getCellFormula())

                    .build();

            cheques.add(cheque);

        });

        return  cheques;
    }


    public List<?> toList(Iterator<?> iterator){
        return IteratorUtils.toList(iterator);
    }


    public void imprimir(List<Cheque> cheques){
        cheques.forEach(System.out::println);
    }

}
