
package com.mycompany.conversorexceltelefonico;
import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.BufferedWriter;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class JanelaInicial extends javax.swing.JPanel {

    private File arquivoExcelSelecionado;
    private File diretorioSalvoSelecionado;
    
    public JanelaInicial() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        lblSelecionarExcel = new javax.swing.JLabel();
        btnSelecionaExcel = new javax.swing.JButton();
        lblSelecionaLocalSalvar = new javax.swing.JLabel();
        btnSalvaNovoExcel = new javax.swing.JButton();
        btnProcessarESalvarExcel = new javax.swing.JButton();

        setBorder(javax.swing.BorderFactory.createTitledBorder("Conversor Excel Relatorio 3CX"));

        lblSelecionarExcel.setText("Selecione o excel a ser trabalhado pelo software");

        btnSelecionaExcel.setText("Selecionar");
        btnSelecionaExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSelecionaExcelActionPerformed(evt);
            }
        });

        lblSelecionaLocalSalvar.setText("Selecioner aonde deseja salvar o arquivo");

        btnSalvaNovoExcel.setText("Selecionar");
        btnSalvaNovoExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSalvaNovoExcelActionPerformed(evt);
            }
        });

        btnProcessarESalvarExcel.setText("Executar");
        btnProcessarESalvarExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnProcessarESalvarExcelActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(42, 42, 42)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btnSalvaNovoExcel)
                    .addComponent(lblSelecionaLocalSalvar)
                    .addComponent(btnSelecionaExcel)
                    .addComponent(lblSelecionarExcel))
                .addContainerGap(97, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(btnProcessarESalvarExcel)
                .addGap(78, 78, 78))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(44, 44, 44)
                .addComponent(lblSelecionarExcel)
                .addGap(18, 18, 18)
                .addComponent(btnSelecionaExcel)
                .addGap(24, 24, 24)
                .addComponent(lblSelecionaLocalSalvar)
                .addGap(18, 18, 18)
                .addComponent(btnSalvaNovoExcel)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 33, Short.MAX_VALUE)
                .addComponent(btnProcessarESalvarExcel)
                .addGap(39, 39, 39))
        );
    }// </editor-fold>//GEN-END:initComponents

    private void btnSelecionaExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSelecionaExcelActionPerformed
        JFileChooser folderChooser = new JFileChooser();
        folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = folderChooser.showSaveDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            diretorioSalvoSelecionado = folderChooser.getSelectedFile();
            JOptionPane.showMessageDialog(this, "Diretório selecionado: " + diretorioSalvoSelecionado.getAbsolutePath());
            processarESalvarExcel(arquivoExcelSelecionado, new File(diretorioSalvoSelecionado, "RelatorioProcessado.csv"));
        }
    }//GEN-LAST:event_btnSelecionaExcelActionPerformed

    private void btnSalvaNovoExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSalvaNovoExcelActionPerformed
        JFileChooser folderChooser = new JFileChooser();
        folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = folderChooser.showSaveDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            diretorioSalvoSelecionado = folderChooser.getSelectedFile();
            JOptionPane.showMessageDialog(this, "Diretório selecionado: " + diretorioSalvoSelecionado.getAbsolutePath());
            processarESalvarExcel(arquivoExcelSelecionado, new File(diretorioSalvoSelecionado, "RelatorioProcessado.csv"));
        }
    }//GEN-LAST:event_btnSalvaNovoExcelActionPerformed

    private void btnProcessarESalvarExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnProcessarESalvarExcelActionPerformed
       
    }//GEN-LAST:event_btnProcessarESalvarExcelActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnProcessarESalvarExcel;
    private javax.swing.JButton btnSalvaNovoExcel;
    private javax.swing.JButton btnSelecionaExcel;
    private javax.swing.JLabel lblSelecionaLocalSalvar;
    private javax.swing.JLabel lblSelecionarExcel;
    // End of variables declaration//GEN-END:variables
    
    private void processarESalvarExcel(File excelFile, File csvFile) {
        try (FileInputStream arquivoExcel = new FileInputStream(excelFile);
             Workbook workbook = WorkbookFactory.create(arquivoExcel);
             BufferedWriter writer = new BufferedWriter(new FileWriter(csvFile))) {
            Sheet sheet = workbook.getSheetAt(0);
            writer.write("Data,Hora,ID Chamador,Destino,Falando\n");
            for (Row row : sheet) {
                // Processamento das células, conforme a lógica anterior...
                // Adapte este trecho conforme necessário para ler e escrever os dados corretos.
            }
            JOptionPane.showMessageDialog(this, "Arquivo CSV salvo em: " + csvFile.getAbsolutePath());
        } catch (IOException e) {
            JOptionPane.showMessageDialog(this, "Erro ao processar o arquivo: " + e.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
        }
    }
}

