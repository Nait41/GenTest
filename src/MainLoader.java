import data.ExceptionList;
import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.*;
import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

public class MainLoader extends JFrame {
    XWPFDocument workbook;
    XWPFDocument workbookTemp;
    SortedTable sortedTable;
    String nameObr;
    public MainLoader(String nameObr) throws IOException, InvalidFormatException {
        File file = new File("C:\\Program Files\\gentest_obr\\" + nameObr + ".docx");
        workbook = new XWPFDocument(new FileInputStream(file));
        if (MainController.mediumRangeOption){
            File fileException = new File("C:\\Program Files\\gentest_obr\\exceptionCheckObrFile\\" + nameObr + ".docx");
            workbookTemp = new XWPFDocument(new FileInputStream(fileException));
        }
        this.nameObr = nameObr;
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void setFileNameForFifth(InfoList infoList){
        XWPFRun run = workbook.getTables().get(0).getRow(12).getCell(0).getParagraphs().get(0).createRun();
        run.setFontSize(11);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.fileName.replace(".xlsx", ""));
    }

    public void setFileNameForFirst(InfoList infoList){
        XWPFRun run = workbook.getTables().get(0).getRow(11).getCell(0).getParagraphs().get(0).createRun();
        run.setFontSize(11);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.fileName.replace(".xlsx", ""));
    }

    public void setFileNameForSecond(InfoList infoList){
        XWPFRun run = workbook.getTables().get(0).getRow(0).getCell(1).getParagraphs().get(3).createRun();
        run.setFontSize(10);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.fileName.replace(".xlsx", ""));
    }
    public void setFileNameForThird(InfoList infoList){
        XWPFRun run = workbook.getTables().get(0).getRow(0).getCell(1).getParagraphs().get(4).createRun();
        run.setFontSize(10);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.fileName.replace(".xlsx", ""));
    }

    public void setFourTableFormatForSecond(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        for(int i = 1; i < workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRows().size();i++){
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                        || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                        && genusSpecies.get(j).get(0).contains("/")
                        && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                        || genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().
                        replace(" ", "_"))
                        || genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().
                        replace("_", " "))
                )
                {
                    XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText(genusSpecies.get(j).get(1));
                    for(int k = 0; k<infoList.algs.size();k++)
                    {
                        if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                                && infoList.algs.get(k).get(0).contains("/")
                                && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                || workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                .replace("_", " "))
                                || workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                .replace(" ", "_"))
                        )
                        {
                            if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).getText().equals("0.0"))
                            {
                                run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Century Gothic");
                                run.setText(infoList.algs.get(k).get(2));
                                break;
                            }
                            else{
                                if(!infoList.algs.get(k).get(1).equals("0.0"))
                                {
                                    if(checkValueRange(infoList.algs.get(k).get(1), workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).getText())){
                                        run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(2));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (MainController.mediumRangeOption)
                {
                    if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getTableCells().size() == 4) {
                        if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")) {
                            for(int k = 0; k<infoList.algs.size();k++) {
                                if ((workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                        || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                                        && infoList.algs.get(k).get(0).contains("/")
                                        && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")))
                                        && workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")
                                ) {
                                    if(infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")){
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(1));
                                        run = workbookTemp.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(1));
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getTableCells().size() == 4)
            {
                if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("не определен");
                    ExceptionList.exceptBact.add(new ArrayList<>());
                    ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText());
                    MainController.exceptCheck = true;
                }
                if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("0.0");
                }
                if(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("отсутствует/крайне низкое/не идентифицирован");
                }
            }
        }
    }

    public void setAdditionForThird(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        int parNumber = 0;
        int counter = 0;
        for(int i = 0; i < workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().size();i++)
        {
            if(workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(i).getText().equals("Дополнение") || workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(i).getText().contains("ДОПОЛНЕНИЕ")){
                parNumber = i;
                break;
            }
        }

        CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
        cTAbstractNum.setAbstractNumId(BigInteger.valueOf(30));

        CTLvl cTLvl = cTAbstractNum.addNewLvl();
        cTLvl.setIlvl(BigInteger.valueOf(0));
        cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
        cTLvl.addNewLvlText().setVal("%1.");
        cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
        cTLvl.addNewRPr();
        cTLvl.getRPr().addNewSz().setVal(9*2);
        cTLvl.getRPr().addNewSzCs().setVal(9*2);
        cTLvl.getRPr().addNewRFonts().setAscii("Century Gothic");

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
        XWPFNumbering numbering = workbook.createNumbering();
        BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
        BigInteger numID = numbering.addNum(abstractNumID);

        parNumber+=1;
        boolean checkBacter = true;
        for(int d = 0; d < infoList.uniqBact.size(); d++)
        {
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                for(int k = 0; k < infoList.algs.size();k++)
                {
                    if(infoList.algs.get(k).size() == 5)
                    {
                        if(infoList.uniqBact.get(d).equals(infoList.algs.get(k).get(3)))
                        {
                            if(infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0))){
                                if(infoList.algs.get(k).get(1).equals("0.0")){
                                    if(genusSpecies.get(j).get(1).equals("0.0")){
                                        if(checkBacter)
                                        {
                                            XmlCursor xmlCursor = workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                            XWPFParagraph xwpfParagraph = workbook.getTables().get(0).getRow(1).getCell(0).insertNewParagraph(xmlCursor);
                                            xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                            xwpfParagraph.setIndentationLeft(0);
                                            XWPFRun run = xwpfParagraph.createRun();
                                            run.setFontSize(9);
                                            run.setBold(true);
                                            run.setItalic(true);
                                            run.setUnderline(UnderlinePatterns.SINGLE);
                                            run.setFontFamily("Century Gothic");
                                            run.setText(infoList.uniqBact.get(d));
                                            run.addBreak();
                                            checkBacter = false;
                                            counter++;
                                        }
                                        XmlCursor xmlCursor = workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                        XWPFParagraph xwpfParagraph = workbook.getTables().get(0).getRow(1).getCell(0).insertNewParagraph(xmlCursor);
                                        xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                        xwpfParagraph.setIndentationLeft(0);
                                        xwpfParagraph.setNumID(numID);
                                        XWPFRun run = xwpfParagraph.createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(4));
                                        run.addBreak();
                                        counter++;
                                        break;
                                    }
                                } else {
                                    if(checkValueRange(infoList.algs.get(k).get(1), genusSpecies.get(j).get(1))){
                                        if(checkBacter)
                                        {
                                            XmlCursor xmlCursor = workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                            XWPFParagraph xwpfParagraph = workbook.getTables().get(0).getRow(1).getCell(0).insertNewParagraph(xmlCursor);
                                            xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                            xwpfParagraph.setIndentationLeft(0);
                                            XWPFRun run = xwpfParagraph.createRun();
                                            run.setFontSize(9);
                                            run.setBold(true);
                                            run.setItalic(true);
                                            run.setUnderline(UnderlinePatterns.SINGLE);
                                            run.setFontFamily("Century Gothic");
                                            run.setText(infoList.uniqBact.get(d));
                                            run.addBreak();
                                            counter++;
                                            checkBacter = false;
                                        }
                                        XmlCursor xmlCursor = workbook.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                        XWPFParagraph xwpfParagraph = workbook.getTables().get(0).getRow(1).getCell(0).insertNewParagraph(xmlCursor);
                                        xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                        xwpfParagraph.setIndentationLeft(0);
                                        xwpfParagraph.setNumID(numID);
                                        XWPFRun run = xwpfParagraph.createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(4));
                                        run.addBreak();
                                        counter++;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            checkBacter = true;
        }
    }

    public void setPhylum(InfoList infoList){
        for(int i = 0; i < workbook.getTables().get(2).getRows().size();i++){
            for (int j = 0; j < infoList.phylum.size(); j++)
            {
                if(workbook.getTables().get(2).getRow(i).getCell(0).getText().equals(infoList.phylum.get(j).get(0)))
                {
                    XWPFRun run = workbook.getTables().get(2).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText(infoList.phylum.get(j).get(1));

                }
            }
            if(workbook.getTables().get(2).getRow(i).getCell(1).getText().equals("")){
                XWPFRun run = workbook.getTables().get(2).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                run.setFontSize(9);
                run.setFontFamily("Century Gothic");
                run.setText("0.0");
            }
        }
    }

    public void setRatio(InfoList infoList){
        int i = 0;
        double bact = 0, firm = 0, acti = 0, prot = 0;
        for (int j = 0; j < infoList.phylum.size(); j++)
        {
            if(infoList.phylum.get(j).get(0).equals("Bacteroidota"))
            {
                bact = Double.parseDouble(infoList.phylum.get(j).get(1).replace(",", "."));
            }
            if(infoList.phylum.get(j).get(0).equals("Firmicutes"))
            {
                firm = Double.parseDouble(infoList.phylum.get(j).get(1).replace(",", "."));
            }
            if(infoList.phylum.get(j).get(0).equals("Proteobacteria"))
            {
                prot = Double.parseDouble(infoList.phylum.get(j).get(1).replace(",", "."));
            }
            if(infoList.phylum.get(j).get(0).equals("Actinobacteriota"))
            {
                acti = Double.parseDouble(infoList.phylum.get(j).get(1).replace(",", "."));
            }
        }
        XWPFRun run = workbook.getTables().get(3).getRow(1).getCell(1).getParagraphs().get(0).createRun();
        run.setFontSize(9);
        run.setFontFamily("Century Gothic");
        run.setText(String.format("%(.2f",(bact/firm)));
        run = workbook.getTables().get(3).getRow(2).getCell(1).getParagraphs().get(0).createRun();
        run.setFontSize(9);
        run.setFontFamily("Century Gothic");
        run.setText(String.format("%(.2f",(firm/prot)));
        run = workbook.getTables().get(3).getRow(3).getCell(1).getParagraphs().get(0).createRun();
        run.setFontSize(9);
        run.setFontFamily("Century Gothic");
        run.setText(String.format("%(.2f",(firm/acti)));
    }

    public void setFiveFormat(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        for(int i = 1; i < workbook.getTables().get(numberTable).getRows().size();i++){
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                        || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                        && genusSpecies.get(j).get(0).contains("/")
                        && !(genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                ))
                {
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText(genusSpecies.get(j).get(1));
                    for(int k = 0; k<infoList.algs.size();k++)
                    {
                        if(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                && infoList.algs.get(k).get(0).contains("/")
                                && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                        )
                        {
                            if(workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).getText().equals("0.0"))
                            {
                                run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Century Gothic");
                                run.setText(infoList.algs.get(k).get(2));
                                break;
                            }
                            else{
                                if(!infoList.algs.get(k).get(1).equals("0.0"))
                                {
                                    if(checkValueRange(infoList.algs.get(k).get(1), workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).getText())){
                                        run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(2));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (MainController.mediumRangeOption) {
                    if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5) {
                        if (workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")) {
                            for(int k = 0; k<infoList.algs.size();k++) {
                                if ((workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                        || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                        && infoList.algs.get(k).get(0).contains("/")
                                        && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                        && workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals(""))
                                ) {
                                    if(infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")){
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(1));
                                        run = workbookTemp.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(1));
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("не определен");
                    ExceptionList.exceptBact.add(new ArrayList<>());
                    ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                    MainController.exceptCheck = true;
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("0.0");
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(4).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("отсутствует/крайне низкое/не идентифицирован");
                }
            }
        }
    }

    public void setFourFormat(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        for(int i = 1; i < workbook.getTables().get(numberTable).getRows().size();i++){
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 4)
                {
                    if(genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace(" ", "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace("_", " "))
                    )
                    {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText(genusSpecies.get(j).get(1));
                        for(int k = 0; k<infoList.algs.size();k++)
                        {
                            if(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace("_", " "))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace(" ", "_"))
                            )
                            {
                                if(workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).getText().equals("0.0"))
                                {
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(2));
                                    break;
                                }
                                else{
                                    if(!infoList.algs.get(k).get(1).equals("0.0"))
                                    {
                                        if(checkValueRange(infoList.algs.get(k).get(1), workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).getText())){
                                            run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Century Gothic");
                                            run.setText(infoList.algs.get(k).get(2));
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
                if (MainController.mediumRangeOption) {
                    if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 4){
                        if(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals(""))
                        {
                            for(int k = 0; k<infoList.algs.size();k++) {
                                if ((workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                        || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                        && infoList.algs.get(k).get(0).contains("/")
                                        && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")))
                                        && workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")
                                ) {
                                    if(infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")){
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(1));
                                        run = workbookTemp.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(1));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 4)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("не определен");
                    ExceptionList.exceptBact.add(new ArrayList<>());
                    ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                    MainController.exceptCheck = true;
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("0.0");
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("отсутствует/крайне низкое/не идентифицирован");
                }
            }
        }
    }

    public void setThreeDoubleFormat(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        boolean checkFirst, checkSecond;
        for(int i = 1; i < workbook.getTables().get(numberTable).getRows().size();i++){
            checkFirst = true;
            checkSecond = true;
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if((genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                        || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                        && genusSpecies.get(j).get(0).contains("/")
                        && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")) && checkFirst
                        && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("")
                ))
                {
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText(genusSpecies.get(j).get(1));
                    checkFirst = false;
                }
                if((workbook.getTables().get(numberTable).getRow(i).getCell(3) != null) && checkSecond)
                {
                    if((genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText())
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText() + "_"))
                    )
                            && !workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")
                    )
                    {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText(genusSpecies.get(j).get(1));
                        checkSecond = false;
                    }
                }
            }
            if (MainController.mediumRangeOption) {
                if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 6){
                    boolean checkFirst_, checkSecond_;
                    for(int k = 0; k<infoList.algs.size();k++) {
                        checkFirst_ = true;
                        checkSecond_ = true;
                        if(!workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(""))
                        {
                            if (workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    && checkFirst_ && workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")
                            ) {
                                if(infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")){
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(1));
                                    run = workbookTemp.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(1));
                                    checkFirst_ = false;
                                }
                            }
                        }
                        if(!workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals(""))
                        {
                            if (workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText() + "_"))
                                    && checkSecond_ && workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")
                            ) {
                                if(infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")){
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(1));
                                    run = workbookTemp.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(1));
                                    checkSecond_ = false;
                                }
                            }
                        }
                    }
                }
            }
            if (workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 6) {
                if (!workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("")) {
                    if (workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")) {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("не определен");
                        ExceptionList.exceptBact.add(new ArrayList<>());
                        ExceptionList.exceptBact.get(ExceptionList.exceptBact.size() - 1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                        MainController.exceptCheck = true;
                    }
                    if (workbook.getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")) {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("0.0");
                    }
                }
                if (!workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")) {
                    if (!workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")) {
                        if (workbook.getTables().get(numberTable).getRow(i).getCell(4).getText().equals("")) {
                            XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Century Gothic");
                            run.setText("не определен");
                            ExceptionList.exceptBact.add(new ArrayList<>());
                            ExceptionList.exceptBact.get(ExceptionList.exceptBact.size() - 1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                            MainController.exceptCheck = true;
                        }
                        if (workbook.getTables().get(numberTable).getRow(i).getCell(5).getText().equals("")) {
                            XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Century Gothic");
                            run.setText("0.0");
                        }
                    }
                }
            }
        }
    }

    public void setTwoFormat(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        for(int i = 0; i < workbook.getTables().get(numberTable).getRows().size();i++){
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(genusSpecies.get(j).get(0)))
                {
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText(genusSpecies.get(j).get(1));

                }
            }
        }
    }

    public void setAddition(InfoList infoList){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        int parNumber = 0;
        int counter = 0;
        for(int i = 0; i < workbook.getParagraphs().size();i++)
        {
            if(workbook.getParagraphs().get(i).getText().equals("Дополнение") || workbook.getParagraphs().get(i).getText().contains("ДОПОЛНЕНИЕ")){
                parNumber = i;
                break;
            }
        }

        CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
        cTAbstractNum.setAbstractNumId(BigInteger.valueOf(30));

        CTLvl cTLvl = cTAbstractNum.addNewLvl();
        cTLvl.setIlvl(BigInteger.valueOf(0));
        cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
        cTLvl.addNewLvlText().setVal("%1.");
        cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
        cTLvl.addNewRPr();
        cTLvl.getRPr().addNewSz().setVal(9*2);
        cTLvl.getRPr().addNewSzCs().setVal(9*2);
        cTLvl.getRPr().addNewRFonts().setAscii("Century Gothic");

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
        XWPFNumbering numbering = workbook.createNumbering();
        BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
        BigInteger numID = numbering.addNum(abstractNumID);

        parNumber+=1;
        boolean checkBacter = true;
        for(int d = 0; d < infoList.uniqBact.size(); d++)
        {
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                for(int k = 0; k < infoList.algs.size();k++)
                {
                    if(infoList.algs.get(k).size() == 5)
                    {
                        if(infoList.uniqBact.get(d).equals(infoList.algs.get(k).get(3)))
                        {
                            if(infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0))){
                                if(infoList.algs.get(k).get(1).equals("0.0")){
                                    if(genusSpecies.get(j).get(1).equals("0.0")){
                                        if(checkBacter)
                                        {
                                            XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                            XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                            xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                            xwpfParagraph.setIndentationLeft(0);
                                            XWPFRun run = xwpfParagraph.createRun();
                                            run.setFontSize(9);
                                            run.setBold(true);
                                            run.setItalic(true);
                                            run.setUnderline(UnderlinePatterns.SINGLE);
                                            run.setFontFamily("Century Gothic");
                                            run.setText(infoList.uniqBact.get(d));
                                            run.addBreak();
                                            checkBacter = false;
                                            counter++;
                                        }
                                        XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                        XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                        xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                        xwpfParagraph.setIndentationLeft(0);
                                        xwpfParagraph.setNumID(numID);
                                        XWPFRun run = xwpfParagraph.createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(4));
                                        run.addBreak();
                                        counter++;
                                        break;
                                    }
                                } else {
                                    if(checkValueRange(infoList.algs.get(k).get(1), genusSpecies.get(j).get(1))){
                                        if(checkBacter)
                                        {
                                            XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                            XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                            xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                            xwpfParagraph.setIndentationLeft(0);
                                            XWPFRun run = xwpfParagraph.createRun();
                                            run.setFontSize(9);
                                            run.setBold(true);
                                            run.setItalic(true);
                                            run.setUnderline(UnderlinePatterns.SINGLE);
                                            run.setFontFamily("Century Gothic");
                                            run.setText(infoList.uniqBact.get(d));
                                            run.addBreak();
                                            counter++;
                                            checkBacter = false;
                                        }
                                        XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                        XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                        xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                        xwpfParagraph.setIndentationLeft(0);
                                        xwpfParagraph.setNumID(numID);
                                        XWPFRun run = xwpfParagraph.createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(infoList.algs.get(k).get(4));
                                        run.addBreak();
                                        counter++;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            checkBacter = true;
        }
    }

    public void setTwoFormatWithSer(InfoList infoList, int numberTable, String numberSer) throws IOException, ClassNotFoundException {
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        loadSortedTable(numberSer);
        ArrayList<ArrayList<String>> result = new ArrayList<>();
        for (int i = 0 ; i < sortedTable.tableFirst.size(); i++)
        {
            result.add(new ArrayList<>());
            result.get(i).add(sortedTable.tableFirst.get(i));
        }
        for(int i = 0; i < sortedTable.tableFirst.size();i++){
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(result.get(i).get(0).equals(genusSpecies.get(j).get(0)))
                {
                    result.get(i).add(genusSpecies.get(j).get(1));
                    break;
                }
            }
            if(result.get(i).size() == 1){
                result.get(i).add("0.0");
            }
        }
        Collections.sort(result, new Comparator<ArrayList<String>>() {
            @Override
            public int compare(ArrayList<String> o1, ArrayList<String> o2) {
                return Double.compare(Double.parseDouble(o2.get(1)), Double.parseDouble(o1.get(1)));
            }
        });
        for(int i = 0; i < result.size();i++){
            XWPFRun run = workbook.getTables().get(numberTable).getRow(i+1).getCell(0).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setText(result.get(i).get(0));
            run = workbook.getTables().get(numberTable).getRow(i+1).getCell(1).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setText(result.get(i).get(1));
        }
    }

    void loadSortedTable(String nameFileSer) throws IOException, ClassNotFoundException {
        FileInputStream fileInputStream = new FileInputStream("C:\\Program Files\\gentest_obr\\saveSortedTable_" + nameFileSer + ".ser");
        ObjectInputStream objectInputStream = new ObjectInputStream(fileInputStream);
        sortedTable = (SortedTable) objectInputStream.readObject();
    }

    void saveSortedTable(InfoList infoList, int numberTable, String nameFileSer) throws IOException, ClassNotFoundException {
        sortedTable = new SortedTable();
        for(int i = 0; i < workbook.getTables().get(numberTable).getRows().size();i++) {
            System.out.println(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
            if(!workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("Классификация")
                    && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("РОД БАКТЕРИЙ")
                    && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("ВИД БАКТЕРИЙ")
                    && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(""))
            {
                System.out.println(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                sortedTable.tableFirst.add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
            }
        }
        FileOutputStream outputStream = new FileOutputStream("C:\\Program Files\\gentest_obr\\saveSortedTable_" + nameFileSer + ".ser");
        ObjectOutputStream objectOutputStream = new ObjectOutputStream(outputStream);
        objectOutputStream.writeObject(sortedTable);
        objectOutputStream.close();
    }

    boolean checkValueRange(String range, String checkNumber){
        String firstNumber = "", secondNumber = "";
        boolean checkChoice = true;
        for(int i = 0;i<range.length();i++){
            if(checkChoice){
                if(range.charAt(i) == '-' || range.charAt(i) == '–')
                {
                    checkChoice = false;
                }
                else{
                    firstNumber += range.charAt(i);
                }
            }else
            {
                secondNumber += range.charAt(i);
            }
        }

        if(Double.parseDouble(checkNumber) > Double.parseDouble(firstNumber) && Double.parseDouble(checkNumber) < Double.parseDouble(secondNumber)){
            return true;
        }
        else{
            return false;
        }
    }

    public void saveFile(InfoList infoList, File docPath) throws IOException {
        workbook.write(new FileOutputStream(new File(docPath.getPath() + "\\" + infoList.fileName.replace(".xlsx", "")) + ".docx"));
    }

    public void saveObrFile() throws IOException {
        workbookTemp.write(new FileOutputStream(new File("C:\\Program Files\\gentest_obr\\" + nameObr + ".docx")));
        workbookTemp.close();
    }
}

