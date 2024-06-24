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
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.swing.border.TitledBorder;

public class LeitorExcelGUI extends JFrame {
    
    private JPanel painelArquivos;
    private JButton botaoAdicionarArquivo;
    private JButton botaoComparar;
    private JTable tabelaResultados;
    private DefaultTableModel modeloTabela;
    private List<File> arquivosExcel = new ArrayList<>();

    public LeitorExcelGUI() {
        setTitle("Comparador Customizável de Excel");
        setSize(800, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        initComponents();
        createLayout();
        addEventListeners();
    }

    private void initComponents() {
        painelArquivos = new JPanel(new GridLayout(0, 1, 5, 5));
        botaoAdicionarArquivo = new JButton("Adicionar Arquivo Excel");
        botaoComparar = new JButton("Comparar");

        modeloTabela = new DefaultTableModel(new Object[]{"Resultado"}, 0);
        tabelaResultados = new JTable(modeloTabela);
    }

    private void createLayout() {
        JPanel painelSuperior = new JPanel(new BorderLayout());
        painelSuperior.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        painelSuperior.add(botaoAdicionarArquivo, BorderLayout.NORTH);
        painelSuperior.add(new JScrollPane(painelArquivos), BorderLayout.CENTER);
        painelSuperior.add(botaoComparar, BorderLayout.SOUTH);

        add(painelSuperior, BorderLayout.NORTH);
        add(new JScrollPane(tabelaResultados), BorderLayout.CENTER);
    }

    private void addEventListeners() {
        botaoAdicionarArquivo.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                adicionarArquivo();
            }
        });

        botaoComparar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                compararArquivos();
            }
        });
    }

    private void adicionarArquivo() {
        JFileChooser seletorArquivo = new JFileChooser();
        int resultado = seletorArquivo.showOpenDialog(LeitorExcelGUI.this);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            File arquivoSelecionado = seletorArquivo.getSelectedFile();
            arquivosExcel.add(arquivoSelecionado);

            JPanel painelArquivo = new JPanel(new GridLayout(0, 2, 5, 5));
            painelArquivo.setBorder(BorderFactory.createTitledBorder("Arquivo: " + arquivoSelecionado.getName()));

            JTextField campoColunasChave = new JTextField();
            campoColunasChave.setToolTipText("Colunas para chave (ex.: A,B)");
            JTextField campoColunasComparacao = new JTextField();
            campoColunasComparacao.setToolTipText("Colunas para comparação (ex.: C,D)");
            JTextField campoColunaRetorno = new JTextField();
            campoColunaRetorno.setToolTipText("Coluna de retorno (ex.: E)");

            painelArquivo.add(new JLabel("Colunas para Chave:"));
            painelArquivo.add(campoColunasChave);
            painelArquivo.add(new JLabel("Colunas para Comparação:"));
            painelArquivo.add(campoColunasComparacao);
            painelArquivo.add(new JLabel("Coluna de Retorno:"));
            painelArquivo.add(campoColunaRetorno);

            painelArquivos.add(painelArquivo);
            painelArquivos.revalidate();
            painelArquivos.repaint();
        }
    }

    private void compararArquivos() {
        modeloTabela.setRowCount(0);
        Map<String, List<String>> mapaChaves = new HashMap<>();

        for (Component componente : painelArquivos.getComponents()) {
            if (componente instanceof JPanel) {
                JPanel painelArquivo = (JPanel) componente;
                String nomeArquivo = ((TitledBorder) painelArquivo.getBorder()).getTitle().replace("Arquivo: ", "");
                File arquivo = arquivosExcel.stream().filter(f -> f.getName().equals(nomeArquivo)).findFirst().orElse(null);

                if (arquivo != null) {
                    JTextField campoColunasChave = (JTextField) painelArquivo.getComponent(1);
                    JTextField campoColunasComparacao = (JTextField) painelArquivo.getComponent(3);
                    JTextField campoColunaRetorno = (JTextField) painelArquivo.getComponent(5);

                    List<String> colunasChave = Arrays.asList(campoColunasChave.getText().split(","));
                    List<String> colunasComparacao = Arrays.asList(campoColunasComparacao.getText().split(","));
                    String colunaRetorno = campoColunaRetorno.getText();

                    try {
                        Map<String, String> chavesArquivo = lerChavesExcel(arquivo, colunasChave, colunasComparacao, colunaRetorno);
                        for (Map.Entry<String, String> entrada : chavesArquivo.entrySet()) {
                            mapaChaves.computeIfAbsent(entrada.getKey(), k -> new ArrayList<>()).add(entrada.getValue());
                        }
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(LeitorExcelGUI.this, "Erro ao ler o arquivo " + nomeArquivo + ": " + ex.getMessage());
                    }
                }
            }
        }

        for (Map.Entry<String, List<String>> entrada : mapaChaves.entrySet()) {
            modeloTabela.addRow(new Object[]{entrada.getKey() + " -> " + entrada.getValue()});
        }
    }

    private Map<String, String> lerChavesExcel(File arquivo, List<String> colunasChave, List<String> colunasComparacao, String colunaRetorno) throws IOException {
        Map<String, String> chaves = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(arquivo);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            int[] indicesChave = colunasChave.stream().mapToInt(col -> col.charAt(0) - 'A').toArray();
            int[] indicesComparacao = colunasComparacao.stream().mapToInt(col -> col.charAt(0) - 'A').toArray();
            int indiceRetorno = colunaRetorno.charAt(0) - 'A';

            for (Row row : sheet) {
                StringBuilder chave = new StringBuilder();
                boolean chaveValida = true;
                for (int indice : indicesChave) {
                    Cell cell = row.getCell(indice);
                    if (cell == null || cell.toString().isEmpty()) {
                        chaveValida = false;
                        break;
                    }
                    chave.append(cell.toString()).append("-");
                }
                if (chaveValida) {
                    chave.setLength(chave.length() - 1); // Remove o último "-"
                    StringBuilder comparacao = new StringBuilder();
                    for (int indice : indicesComparacao) {
                        Cell cell = row.getCell(indice);
                        comparacao.append(cell != null ? cell.toString() : "").append("-");
                    }
                    comparacao.setLength(comparacao.length() - 1); // Remove o último "-"
                    Cell cellRetorno = row.getCell(indiceRetorno);
                    String retorno = cellRetorno != null ? cellRetorno.toString() : "";
                    if (retorno.isEmpty()) {
                        retorno = "Adicionado";
                    }
                    chaves.put(chave.toString(), comparacao + " -> " + retorno);
                }
            }
        }
        return chaves;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            LeitorExcelGUI gui = new LeitorExcelGUI();
            gui.setVisible(true);
        });
    }
}