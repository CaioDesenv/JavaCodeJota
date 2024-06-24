package com.jotacontabil.processadorxmlgui;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.*;
import java.util.*;

public class ProcessadorXMLGUI extends JFrame {
    private JTable tabela;
    private DefaultTableModel modeloTabela;
    private java.util.List<Map<String, String>> registros = new ArrayList<>();
    private JTextField campoCaminhoPasta;

    public ProcessadorXMLGUI() {
        setTitle("Processador de XML");
        setSize(800, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JPanel painelPrincipal = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();

        campoCaminhoPasta = new JTextField(30);
        JButton botaoProcurar = new JButton("Procurar");
        JButton botaoCarregar = new JButton("Carregar Arquivos XML");

        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        JPanel painelSuperior = new JPanel(new GridBagLayout());
        GridBagConstraints gbcSuperior = new GridBagConstraints();

        gbcSuperior.gridx = 0;
        gbcSuperior.gridy = 0;
        painelSuperior.add(new JLabel("Caminho da Pasta:"), gbcSuperior);

        gbcSuperior.gridx = 1;
        painelSuperior.add(campoCaminhoPasta, gbcSuperior);

        gbcSuperior.gridx = 2;
        painelSuperior.add(botaoProcurar, gbcSuperior);

        gbcSuperior.gridx = 3;
        painelSuperior.add(botaoCarregar, gbcSuperior);

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 4;
        painelPrincipal.add(painelSuperior, gbc);

        String[] nomesColunas = {"Número da Nota", "Data de Emissão", "Valor da Nota", "Valor do ICMS ST"};
        modeloTabela = new DefaultTableModel(nomesColunas, 0);
        tabela = new JTable(modeloTabela);
        JScrollPane scrollPane = new JScrollPane(tabela);

        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 4;
        gbc.fill = GridBagConstraints.BOTH;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        painelPrincipal.add(scrollPane, gbc);

        JButton botaoExportar = new JButton("Exportar para TXT");

        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 4;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 0;
        gbc.weighty = 0;
        painelPrincipal.add(botaoExportar, gbc);

        add(painelPrincipal, BorderLayout.CENTER);

        botaoProcurar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser seletorPasta = new JFileChooser();
                seletorPasta.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int opcao = seletorPasta.showOpenDialog(ProcessadorXMLGUI.this);
                if (opcao == JFileChooser.APPROVE_OPTION) {
                    campoCaminhoPasta.setText(seletorPasta.getSelectedFile().getAbsolutePath());
                }
            }
        });

        botaoCarregar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                carregarArquivosXML(campoCaminhoPasta.getText());
            }
        });

        botaoExportar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                exportarParaArquivo("output.txt");
            }
        });
    }

    private void carregarArquivosXML(String caminhoPasta) {
        registros.clear();
        modeloTabela.setRowCount(0);

        try {
            Files.walk(Paths.get(caminhoPasta))
                .filter(Files::isRegularFile)
                .filter(path -> path.toString().endsWith(".xml"))
                .forEach(path -> registros.add(processarArquivoXML(path.toFile())));

            for (Map<String, String> registro : registros) {
                modeloTabela.addRow(new Object[]{
                    registro.get("Número da Nota"),
                    registro.get("Data de Emissão"),
                    registro.get("Valor da Nota"),
                    registro.get("Valor do ICMS ST")
                });
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Map<String, String> processarArquivoXML(File arquivo) {
        Map<String, String> dados = new HashMap<>();

        try {
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(arquivo);
            doc.getDocumentElement().normalize();

            dados.put("Número da Nota", obterValorTag("nNF", doc));
            dados.put("Data de Emissão", obterValorTag("dhEmi", doc));
            dados.put("Valor da Nota", obterValorTag("vNF", doc));
            dados.put("Valor do ICMS ST", obterValorTag("vST", doc));

        } catch (Exception e) {
            e.printStackTrace();
        }

        return dados;
    }

    private String obterValorTag(String tag, Document doc) {
        NodeList nodeList = doc.getElementsByTagName(tag);
        if (nodeList.getLength() > 0) {
            return nodeList.item(0).getTextContent();
        }
        return "N/A";
    }

    private void exportarParaArquivo(String nomeArquivo) {
        JFileChooser seletorArquivo = new JFileChooser();
        seletorArquivo.setSelectedFile(new File(nomeArquivo));
        int opcao = seletorArquivo.showSaveDialog(this);
        if (opcao == JFileChooser.APPROVE_OPTION) {
            File arquivo = seletorArquivo.getSelectedFile();
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(arquivo))) {
                writer.write("Número da Nota,Data de Emissão,Valor da Nota,Valor do ICMS ST");
                writer.newLine();

                for (Map<String, String> registro : registros) {
                   
                    writer.write(registro.get("Número da Nota") + "," +
                                 registro.get("Data de Emissão") + "," +
                                 registro.get("Valor da Nota") + "," +
                                 registro.get("Valor do ICMS ST"));
                    writer.newLine();
                }
                JOptionPane.showMessageDialog(this, "Arquivo exportado com sucesso!", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException e) {
                JOptionPane.showMessageDialog(this, "Erro ao exportar arquivo: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new ProcessadorXMLGUI().setVisible(true);
            }
        });
    }
}
