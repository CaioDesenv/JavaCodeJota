package model;

public class ConsultaResult {
    private String dataSaida;
    private String numeroNota;
    private String chaveXml;
    private int idEmpresa;

    public ConsultaResult(String dataSaida, String numeroNota, String chaveXml) {
        this.dataSaida = dataSaida;
        this.numeroNota = numeroNota;
        this.chaveXml = chaveXml;
        
    }
    
    public int getIdEmpresa() {
        return idEmpresa;
    }
    
    public String getDataSaida() {
        return dataSaida;
    }

    public String getNumeroNota() {
        return numeroNota;
    }

    public String getChaveXml() {
        return chaveXml;
    }
}
