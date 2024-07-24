package dao;

import model.ConsultaResult;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

public class ConexaoODBCDAO {

    private Connection getConnection() throws Exception {
        Class.forName("sun.jdbc.odbc.JdbcOdbcDriver").newInstance();
        return DriverManager.getConnection("jdbc:odbc:REMJOTA", "REMOTO", "Jota@9960");
    }

    public List<ConsultaResult> executeQuery(String startDate, String endDate, int offset, int limit, int empresaCode) throws Exception {
        List<ConsultaResult> results = new ArrayList<>();
        
        String sql = "SELECT DISTINCT TOP ? " +
                     "a.DATA_SAIDA AS 'Data de Saida', " +
                     "a.nume_sai AS 'Numero da Nota', " +
                     "a.chave_nfe_sai AS 'Chave XML no Domínio' " +
                     "FROM bethadba.efsaidas a " +
                     "WHERE a.codi_emp = ? " +
                     "AND a.dsai_sai BETWEEN ? AND ? " +
                     "AND a.chave_nfe_sai IS NOT NULL " +
                     "AND a.chave_nfe_sai NOT IN (SELECT TOP ? a.chave_nfe_sai " +
                     "FROM bethadba.efsaidas a " +
                     "WHERE a.codi_emp = ? " +
                     "AND a.dsai_sai BETWEEN ? AND ? " +
                     "AND a.chave_nfe_sai IS NOT NULL " +
                     "ORDER BY a.DATA_SAIDA)";

        try (Connection con = getConnection();
             PreparedStatement stmt = con.prepareStatement(sql)) {
            stmt.setInt(1, limit);
            stmt.setInt(2, empresaCode);
            stmt.setString(3, startDate);
            stmt.setString(4, endDate);
            stmt.setInt(5, offset);
            stmt.setInt(6, empresaCode);
            stmt.setString(7, startDate);
            stmt.setString(8, endDate);

            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    String dataSaida = rs.getString("Data de Saida");
                    String numeroNota = rs.getString("Numero da Nota");
                    String chaveXml = rs.getString("Chave XML no Domínio");
                    results.add(new ConsultaResult(dataSaida, numeroNota, chaveXml));
                }
            }
        }

        return results;
    }

    public int countLines(String startDate, String endDate, int empresaCode) throws Exception {
        String sql = "SELECT COUNT(*) AS total " +
                     "FROM bethadba.efsaidas a " +
                     "WHERE a.codi_emp = ? " +
                     "AND a.dsai_sai BETWEEN ? AND ? " +
                     "AND a.chave_nfe_sai IS NOT NULL";

        try (Connection con = getConnection();
             PreparedStatement stmt = con.prepareStatement(sql)) {
            stmt.setInt(1, empresaCode);
            stmt.setString(2, startDate);
            stmt.setString(3, endDate);

            try (ResultSet rs = stmt.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt("total");
                }
            }
        }

        return 0;
    }
}
