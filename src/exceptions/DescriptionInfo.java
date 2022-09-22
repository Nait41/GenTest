package exceptions;

public class DescriptionInfo {
    private String bacteria;
    private String description;
    private String range;
    private String rangeIndex;

    public String getRange() {
        return range;
    }

    public void setRange(String range) {
        this.range = range;
    }

    public String getRangeIndex() {
        return rangeIndex;
    }

    public void setRangeIndex(String rangeIndex) {
        this.rangeIndex = rangeIndex;
    }

    public DescriptionInfo(){}

    public String getBacteria() {
        return bacteria;
    }

    public void setBacteria(String bacteria) {
        this.bacteria = bacteria;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }
}
