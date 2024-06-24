package com.mycompany.conversorexceltelefonico;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.BufferedWriter;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ConversorExcelTelefonico {

    public static void main(String[] args) {
        System.out.println("Data,Hora,ID Chamador,Destino,Falando");
        String caminhoArquivoExcel = "C:\\Users\\jota02\\Downloads\\Relatorio de telefonia.xlsx";
        String caminhoArquivoCSV = "C:\\Users\\jota02\\Downloads\\TesteRelatorioVisaoLogica.csv";
        try {
            FileInputStream arquivoExcel = new FileInputStream(new File(caminhoArquivoExcel));
            Workbook workbook = WorkbookFactory.create(arquivoExcel);
            Sheet sheet = workbook.getSheetAt(0);

            // Abrindo o BufferedWriter para escrever no arquivo CSV
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(caminhoArquivoCSV))) {
                // Escrevendo o cabeçalho no arquivo CSV
                writer.write("Data,Hora,ID Chamador,Destino,Falando\n");

                for (Row row : sheet) {
                    Cell duracaoCell = row.getCell(0);
                    Cell idChamadorCell = row.getCell(1);
                    Cell destinoCell = row.getCell(2);
                    Cell falandoCell = row.getCell(3);

                    if (duracaoCell != null && idChamadorCell != null && destinoCell != null && falandoCell != null) {
                        String duracao = formatarDuracao(duracaoCell.getStringCellValue());
                        String[] dataHora = extrairDataHora(duracaoCell.getStringCellValue());
                        String idChamador = extrairValorEntreParenteses(idChamadorCell.getStringCellValue());
                        String destino = extrairValorEntreParenteses(destinoCell.getStringCellValue());
                        String falando = falandoCell.getStringCellValue();

                        String linhaCSV = dataHora[0] + "," + dataHora[1] + "," + idChamador + "," + destino + "," + falando + "\n";
                        writer.write(linhaCSV);
                    }
                }
            }

            workbook.close();
            arquivoExcel.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String formatarDuracao(String duracao) {
        // Adicione aqui a lógica para formatar a duração da chamada, se necessário
        return duracao;
    }

    public static String[] extrairDataHora(String duracao) {
        Pattern pattern = Pattern.compile("(\\d{2}/\\d{2}/\\d{4}) (\\d{2}:\\d{2}:\\d{2})");
        Matcher matcher = pattern.matcher(duracao);
        String[] dataHora = new String[2];
        if (matcher.find()) {
            dataHora[0] = matcher.group(1);
            dataHora[1] = matcher.group(2);
        }
        return dataHora;
    }

    public static String extrairValorEntreParenteses(String texto) {
    // Primeiro, tenta encontrar números dentro dos parênteses.
    Pattern pattern = Pattern.compile("\\((.*?)\\)");
    Matcher matcher = pattern.matcher(texto);
    if (matcher.find()) {
        return matcher.group(1); // Retorna o valor encontrado dentro dos parênteses.
    } else {
        // Se não encontrar nada dentro dos parênteses, tenta encontrar números diretamente.
        pattern = Pattern.compile("\\d+");
        matcher = pattern.matcher(texto);
        if (matcher.find()) {
            return matcher.group(0); // Retorna o número encontrado.
        }
    }
    return null; // Retorna nulo se não encontrar nenhum padrão correspondente.
}

}