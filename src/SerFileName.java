public enum SerFileName {
    FIRST ("First"),
    SECOND ("Second"),
    THIRD ("Third"),
    FOURTH ("Fourth"),
    FIFTH ("Fifth"),
    SIXTH ("Sixth"),
    SEVENTH ("Seventh"),
    EIGHTH ("Eighth"),
    NINTH ("Ninth"),
    TENTH ("Tenth");
    private String title;
    SerFileName(String title){
        this.title = title;
    }
    public String toString() {
        return title;
    }
}