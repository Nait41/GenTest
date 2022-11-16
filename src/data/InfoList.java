package data;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

public class InfoList {
    public String fileName = "load";
    public ArrayList<ArrayList<String>> phylum = new ArrayList<>();
    public ArrayList<ArrayList<String>> genus = new ArrayList<>();
    public ArrayList<ArrayList<String>> species = new ArrayList<>();
    public ArrayList<ArrayList<String>> family = new ArrayList<>();
    public ArrayList<ArrayList<String>> algs = new ArrayList<>();
    public ArrayList<String> uniqBact = new ArrayList<>();
    public ArrayList<String> bioIndex = new ArrayList<>();
}
