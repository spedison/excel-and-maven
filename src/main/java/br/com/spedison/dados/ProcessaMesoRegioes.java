package br.com.spedison.dados;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintStream;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.TreeSet;

@Data
@EqualsAndHashCode(of = "nome")
class Mesoregiao {
    private String nome;
    private TreeSet<Microregiao> microregioes = new TreeSet<Microregiao>();
}

@Data
@EqualsAndHashCode(of = "nome")
class Microregiao {
    private String nome;
    private TreeSet<String> municipio = new TreeSet<String>();
}

@Data
@EqualsAndHashCode(of = "nome")
class Municipio {
    private String nome;
    private Double rendaTodos;
    private Double rendaEmpregados;
    private Double rendaCarteiraTrabalho;
    private Double rendaMilitares;
    private Double rendaSemCarteiraTrab;
    private Double rendaPropria;
    private Double rendaEmpregadores;
}

@Data
@EqualsAndHashCode(of = "nome")
@AllArgsConstructor
@ToString(includeFieldNames = true)
class Linha {
    private Long id;
    private String Mesoregiao;
    private String microregiao;
    private String municipio;
    private Double rendaTodos;
    private Double rendaEmpregados;
    private Double rendaCarteiraTrabalho;
    private Double rendaMilitares;
    private Double rendaSemCarteiraTrab;
    private Double rendaPropria;
    private Double rendaEmpregadores;
}


public class ProcessaMesoRegioes {

    List<Linha> linhas = new LinkedList<Linha>();
    TreeSet<String> mesoregioes = new TreeSet<String>();
    TreeSet<String> microregioes = new TreeSet<String>();

    void adicionaDados(Long id,
                       String mesoregiao,
                       String microregiao,
                       String municipio,
                       Double rendaTodos,
                       Double rendaEmpregados,
                       Double rendaCarteiraTrabalho,
                       Double rendaMilitares,
                       Double rendaSemCarteiraTrab,
                       Double rendaPropria,
                       Double rendaEmpregadores) {

        Linha linha = new Linha(id,
                mesoregiao,
                microregiao,
                municipio,
                rendaTodos,
                rendaEmpregados,
                rendaCarteiraTrabalho,
                rendaMilitares,
                rendaSemCarteiraTrab,
                rendaPropria,
                rendaEmpregadores);

        linhas.add(linha);

        mesoregioes.add(mesoregiao);
        microregioes.add(microregiao);
    }

    public void carregaExcel(String fileName) throws Exception {
        File myFile = new File(fileName);
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
        rowIterator.next(); // Pula o cabeçalho

        Long rowId = 0L;

        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {

            Row row = rowIterator.next(); // Dados !

            int cell = 0;
            // For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();

            Long id = rowId++;

            Cell cel = cellIterator.next();
            String Mesoregiao = cel.getStringCellValue();

            if(Mesoregiao == null || Mesoregiao.trim().isEmpty())
                break;

            cel = cellIterator.next();
            String microregiao = cel.getStringCellValue();

            cel = cellIterator.next();
            String municipio = cel.getStringCellValue();

            cel = cellIterator.next();
            Double rendaTodos = cel.getNumericCellValue();

            cel = cellIterator.next();
            Double rendaEmpregados = cel.getNumericCellValue();

            cel = cellIterator.next();
            Double rendaCarteiraTrabalho = cel.getNumericCellValue();

            cel = cellIterator.next();
            Double rendaMilitares = cel.getNumericCellValue();

            cel = cellIterator.next();
            Double rendaSemCarteiraTrab = cel.getNumericCellValue();

            cel = cellIterator.next();
            Double rendaPropria = cel.getNumericCellValue();

            cel = cellIterator.next();
            Double rendaEmpregadores;
            if(cel.getCellType() == Cell.CELL_TYPE_NUMERIC)
                rendaEmpregadores = cel.getNumericCellValue();
            else
                rendaEmpregadores = null;

            adicionaDados(id,
                    Mesoregiao,
                    microregiao,
                    municipio,
                    rendaTodos,
                    rendaEmpregados,
                    rendaCarteiraTrabalho,
                    rendaMilitares,
                    rendaSemCarteiraTrab,
                    rendaPropria,
                    rendaEmpregadores);

        }

        PrintStream a = System.out;

        linhas.stream().limit(10).forEach(x -> a.println(x));

        mesoregioes.stream().limit(10).forEach(x -> a.println(x));

        microregioes.stream().limit(10).forEach(x -> a.println(x));




    }


    static public void main(String ... args) throws  Exception{
        ProcessaMesoRegioes pmr = new ProcessaMesoRegioes();
        pmr.carregaExcel("C:\\Users\\spedi\\OneDrive\\Cursos\\Base-Dados-Alagoas.xlsx");
    }

    /*

            // Return first sheet from the XLSX workbook




                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t");
                        break;
                    default :

                    }
                }
                System.out.println("");
            }
    * */

}
