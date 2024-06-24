
package com.jotacontabil.formatadorcontabil;
    
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class LeitorExcelGUI extends JFrame {
    
   private JTextField campoCaminhoArquivo;
    private JTextField campoColunaChave1;
    private JTextField campoColunaChave2;
    private JTextField campoColunaID;
    private JTextField campoColunaClassificacao;
    private JButton botaoProcurar;
    private JButton botaoLer;
    private JButton botaoComparar;
    private JTable tabelaResultados;
    private DefaultTableModel modeloTabela;
    private List<String> chaves = new ArrayList<>();

    public LeitorExcelGUI() {
        setTitle("Leitor de Colunas do Excel");
        setSize(600, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        initComponents();
        createLayout();
        addEventListeners();
    }

    private void initComponents() {
        campoCaminhoArquivo = new JTextField();
        campoColunaChave1 = new JTextField();
        campoColunaChave2 = new JTextField();
        campoColunaID = new JTextField();
        campoColunaClassificacao = new JTextField();
        botaoProcurar = new JButton("Procurar");
        botaoLer = new JButton("Ler Colunas");
        botaoComparar = new JButton("Comparar com outra tabela");

        modeloTabela = new DefaultTableModel(new Object[]{"Chave"}, 0);
        tabelaResultados = new JTable(modeloTabela);
    }

    private void createLayout() {
        JPanel painelSuperior = new JPanel(new GridLayout(6, 2, 5, 5));
        painelSuperior.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        painelSuperior.add(new JLabel("Caminho do Arquivo Excel:"));
        painelSuperior.add(campoCaminhoArquivo);
        painelSuperior.add(botaoProcurar);
        painelSuperior.add(new JLabel("Coluna da Chave 1 (ex.: A):"));
        painelSuperior.add(campoColunaChave1);
        painelSuperior.add(new JLabel("Coluna da Chave 2 (ex.: B):"));
        painelSuperior.add(campoColunaChave2);
        painelSuperior.add(new JLabel("Coluna de ID (ex.: C):"));
        painelSuperior.add(campoColunaID);
        painelSuperior.add(new JLabel("Coluna de Classificação (ex.: D):"));
        painelSuperior.add(campoColunaClassificacao);
        painelSuperior.add(botaoLer);
        painelSuperior.add(botaoComparar);

        add(painelSuperior, BorderLayout.NORTH);
        add(new JScrollPane(tabelaResultados), BorderLayout.CENTER);
    }

    private void addEventListeners() {
        botaoProcurar.addActionListener(new BotaoProcurarListener());
        botaoLer.addActionListener(new BotaoLerListener());
        botaoComparar.addActionListener(new BotaoCompararListener());
    }

    private class BotaoProcurarListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            selecionarArquivo();
        }
    }

    private class BotaoLerListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            lerColunas();
        }
    }

    private class BotaoCompararListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            compararEAtualizarTabela();
        }
    }

    private void selecionarArquivo() {
        JFileChooser seletorArquivo = new JFileChooser();
        int resultado = seletorArquivo.showOpenDialog(LeitorExcelGUI.this);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            File arquivoSelecionado = seletorArquivo.getSelectedFile();
            campoCaminhoArquivo.setText(arquivoSelecionado.getAbsolutePath());
        }
    }

    private void lerColunas() {
        String caminhoArquivo = campoCaminhoArquivo.getText();
        String colunaChave1 = campoColunaChave1.getText().toUpperCase();
        String colunaChave2 = campoColunaChave2.getText().toUpperCase();
        try {
            chaves = lerChavesExcel(caminhoArquivo, colunaChave1, colunaChave2);
            modeloTabela.setRowCount(0);
            for (String chave : chaves) {
                modeloTabela.addRow(new Object[]{chave});
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(LeitorExcelGUI.this, "Erro ao ler o arquivo Excel: " + ex.getMessage());
        }
    }

    private List<String> lerChavesExcel(String caminhoArquivo, String coluna1, String coluna2) throws IOException {
        List<String> chaves = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(caminhoArquivo));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int indiceColuna1 = coluna1.charAt(0) - 'A';
            int indiceColuna2 = coluna2.charAt(0) - 'A';
            for (Row row : sheet) {
                Cell cell1 = row.getCell(indiceColuna1);
                Cell cell2 = row.getCell(indiceColuna2);
                if (cell1 != null && cell2 != null && !cell1.toString().isEmpty() && !cell2.toString().isEmpty()) {
                    String chave = cell1.toString() + " - " + cell2.toString();
                    chaves.add(chave);
                }
            }
        }
        return chaves;
    }

    private void compararEAtualizarTabela() {
        String caminhoArquivo = campoCaminhoArquivo.getText();
        String colunaID = campoColunaID.getText().toUpperCase();
        String colunaClassificacao = campoColunaClassificacao.getText().toUpperCase();
        try {
            atualizarTabela(caminhoArquivo, colunaID, colunaClassificacao);
            JOptionPane.showMessageDialog(LeitorExcelGUI.this, "Comparação e atualização concluídas.");
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(LeitorExcelGUI.this, "Erro ao comparar e atualizar a tabela: " + ex.getMessage());
        }
    }

    private void atualizarTabela(String caminhoArquivo, String colunaID, String colunaClassificacao) throws IOException {
        File arquivo = new File(caminhoArquivo);
        try (FileInputStream fis = new FileInputStream(arquivo);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<String, String> mapaChaves = criarMapaChaves();

            int indiceColunaID = colunaID.charAt(0) - 'A';
            int indiceColunaClassificacao = colunaClassificacao.charAt(0) - 'A';

            for (Row row : sheet) {
                Cell cellID = row.getCell(indiceColunaID);
                Cell cellClassificacao = row.getCell(indiceColunaClassificacao);

                if (cellID != null) {
                    String id = cellID.toString();
                    if (cellClassificacao == null || cellClassificacao.toString().isEmpty()) {
                        String classificacao = mapaChaves.get(id);
                        if (classificacao != null) {
                            if (cellClassificacao == null) {
                                cellClassificacao = row.createCell(indiceColunaClassificacao);
                            }
                            cellClassificacao.setCellValue(classificacao);
                        }
                    }
                }
            }

            salvarAlteracoes(workbook, arquivo);
        }
    }

    private Map<String, String> criarMapaChaves() {
        Map<String, String> mapaChaves = new HashMap<>();
        for (String chave : chaves) {
            String[] partesChave = chave.split(" - ");
            if (partesChave.length == 2) {
                mapaChaves.put(partesChave[0], partesChave[1]);
            }
        }
        return mapaChaves;
    }

    private void salvarAlteracoes(Workbook workbook, File arquivo) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(arquivo)) {
            workbook.write(fos);
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            LeitorExcelGUI gui = new LeitorExcelGUI();
            gui.setVisible(true);
        });
    }
}