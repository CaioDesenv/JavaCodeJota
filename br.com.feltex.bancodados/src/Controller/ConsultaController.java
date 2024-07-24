package controller;

import dao.ConexaoODBCDAO;
import model.ConsultaResult;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.List;

public class ConsultaController extends JFrame {
    private JTextField startDateField;
    private JTextField endDateField;
    private JTextField empresaCodeField;
    private DefaultTableModel tableModel;
    private ConexaoODBCDAO dao;

    public ConsultaController() {
        dao = new ConexaoODBCDAO();

        setTitle("Consulta de Dados");
        setSize(1200, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        // Painel para inserção de datas e código da empresa
        JPanel inputPanel = new JPanel();
        inputPanel.add(new JLabel("Código da Empresa:"));
        empresaCodeField = new JTextField(5);
        inputPanel.add(empresaCodeField);

        inputPanel.add(new JLabel("Data Inicial (YYYY-MM-DD):"));
        startDateField = new JTextField(10);
        inputPanel.add(startDateField);

        inputPanel.add(new JLabel("Data Final (YYYY-MM-DD):"));
        endDateField = new JTextField(10);
        inputPanel.add(endDateField);

        JButton executeButton = new JButton("Executar Consulta");
        inputPanel.add(executeButton);
        JButton countButton = new JButton("Contar Linhas");
        inputPanel.add(countButton);
        JButton copyButton = new JButton("Copiar Linhas");
        inputPanel.add(copyButton);

        add(inputPanel, BorderLayout.NORTH);

        // Tabela para exibição dos resultados
        tableModel = new DefaultTableModel();
        JTable table = new JTable(tableModel);
        add(new JScrollPane(table), BorderLayout.CENTER);

        // Listener para o botão de executar consulta
        executeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                executeQuery(0, 700);
            }
        });

        // Listener para o botão de contar linhas
        countButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                countLines();
            }
        });

        // Listener para o botão de copiar linhas
        copyButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                copyLines();
            }
        });
    }

    private void executeQuery(int offset, int limit) {
        String startDate = startDateField.getText();
        String endDate = endDateField.getText();
        int empresaCode = Integer.parseInt(empresaCodeField.getText());

        try {
            List<ConsultaResult> results = dao.executeQuery(startDate, endDate, offset, limit, empresaCode);
            tableModel.setRowCount(0);
            tableModel.setColumnCount(0);

            // Adicionar colunas à tabela
            tableModel.addColumn("Data de Saida");
            tableModel.addColumn("Numero da Nota");
            tableModel.addColumn("Chave XML no Domínio");

            // Adicionar linhas à tabela
            for (ConsultaResult result : results) {
                tableModel.addRow(new Object[]{
                        result.getDataSaida(), result.getNumeroNota(), result.getChaveXml()
                });
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void countLines() {
        String startDate = startDateField.getText();
        String endDate = endDateField.getText();
        int empresaCode = Integer.parseInt(empresaCodeField.getText());

        try {
            int totalLines = dao.countLines(startDate, endDate, empresaCode);
            if (totalLines > 700) {
                JOptionPane.showMessageDialog(this, "Consulta maior que 700 linhas");
            } else {
                JOptionPane.showMessageDialog(this, "Total de linhas: " + totalLines);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void copyLines() {
        int offset = 0;
        boolean hasMore = true;

        while (hasMore) {
            executeQuery(offset, 700);

            int totalCopiedLines = tableModel.getRowCount();
            copyToClipboard(totalCopiedLines);

            if (totalCopiedLines >= 700) {
                int response = JOptionPane.showConfirmDialog(this, "Deseja contar o restante da consulta?");
                if (response != JOptionPane.YES_OPTION) {
                    hasMore = false;
                } else {
                    offset += 700;
                }
            } else {
                JOptionPane.showMessageDialog(this, "Consulta com menos de 700 linhas");
                hasMore = false;
            }
        }
    }

    private void copyToClipboard(int linesToCopy) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < linesToCopy; i++) {
            sb.append(tableModel.getValueAt(i, 0)).append("\t")
              .append(tableModel.getValueAt(i, 1)).append("\t")
              .append(tableModel.getValueAt(i, 2)).append("\n");
        }

        StringSelection selection = new StringSelection(sb.toString());
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(selection, selection);
        JOptionPane.showMessageDialog(this, linesToCopy + " linhas copiadas para a área de transferência");
    }

    public static void main(String[] args) {
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                new ConsultaController().setVisible(true);
            }
        });
    }
}
