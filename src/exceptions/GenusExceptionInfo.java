package exceptions;

public class GenusExceptionInfo {
    private String bacteria;
    private String genus;

    public GenusExceptionInfo(){}

    public String getBacteria() {
        return bacteria;
    }

    public void setBacteria(String bacteria) {
        this.bacteria = bacteria;
    }

    public String getRange() {
        return genus;
    }

    public void setRange(String range) {
        this.genus = range;
    }
}
