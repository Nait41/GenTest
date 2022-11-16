import data.ExceptionList;
import data.InfoList;
import javafx.application.Platform;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.*;
import java.io.*;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

public class MainLoader extends JFrame {
    XWPFDocument workbook;
    XWPFDocument workbookTemp;
    SortedTable sortedTable;
    String nameObr;
    public MainLoader(String nameObr) throws IOException, InvalidFormatException {
        if (!MainController.mediumRangeOption)
        {
            File file = new File(Application.rootDirPath + "\\" + nameObr + ".docx");
            workbook = new XWPFDocument(new FileInputStream(file));
        }
        if (MainController.mediumRangeOption){
            File fileException = new File(Application.rootDirPath + "\\exceptionCheckObrFile\\" + nameObr + ".docx");
            workbookTemp = new XWPFDocument(new FileInputStream(fileException));
            workbook = new XWPFDocument(new FileInputStream(fileException));
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

    public void setBioIndex(InfoList infoList, int numberTable){
        if(nameObr.equals("obr_4")){
            XWPFRun run = workbook.getTables().get(numberTable).getRow(1).getCell(2).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Verdana");
            run.setText(infoList.bioIndex.get(0));
            run = workbook.getTables().get(numberTable).getRow(1).getCell(3).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Verdana");
            run.setText(infoList.bioIndex.get(1));
        }
        if(nameObr.equals("obr_3") || nameObr.equals("obr")){
            XWPFRun run = workbook.getTables().get(numberTable).getRow(1).getCell(2).getParagraphs().get(0).createRun();
            run.setFontSize(12);
            run.setFontFamily("Century Gothic");
            run.setText(infoList.bioIndex.get(0));
            run.setColor("ffffff");
            run.setBold(true);
            run = workbook.getTables().get(numberTable).getRow(1).getCell(3).getParagraphs().get(0).createRun();
            run.setFontSize(12);
            run.setFontFamily("Century Gothic");
            run.setText(infoList.bioIndex.get(1));
            run.setColor("ffffff");
            run.setBold(true);
        }
        if(nameObr.equals("obr_2") || nameObr.equals("obr_1")){
            XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(0).getCell(2).getParagraphs().get(1).createRun();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setText(infoList.bioIndex.get(0));
            run.setColor("0db3b3");
            run.setBold(true);
            run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(0).getCell(3).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setText(infoList.bioIndex.get(1));
            run.setColor("0db3b3");
            run.setBold(true);
        }
    }

    public void setFourTableFormatForSecond(InfoList infoList, int numberTable, MainController mainController){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        double lowLoadStatus;
        double currentStatus = 0;
        boolean checkMediumRange = true;
        lowLoadStatus = 100.0 / workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRows().size();
        for(int i = 1; i < workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRows().size();i++) {
            if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRows().get(i).getTableCells().size() == 4) {
                for (int j = 0; j < genusSpecies.size(); j++) {
                    if (((genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace(" ", "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace("_", " ")))
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")))
                            && !workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals("")
                            && workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")
                    ) {
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText(genusSpecies.get(j).get(1));
                        for (int k = 0; k < infoList.algs.size(); k++) {
                            if ((workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    || workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace("_", " "))
                                    || workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace(" ", "_"))
                            ) && !workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals("")
                            ) {
                                if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).getText().equals("0.0")) {
                                    run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(2));
                                    break;
                                } else {
                                    if (!infoList.algs.get(k).get(1).equals("0.0") && checkMediumRange) {
                                        if (checkValueRange(infoList.algs.get(k).get(1), workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).getText())) {
                                            checkMediumRange = false;
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
                    if (MainController.mediumRangeOption) {
                        if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getTableCells().size() == 4) {
                            if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")) {
                                for (int k = 0; k < infoList.algs.size(); k++) {
                                    if (((workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                            || workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                            .replace("_", " "))
                                            || workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                            .replace(" ", "_")))
                                            || infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())
                                            && infoList.algs.get(k).get(0).contains("/")
                                            && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    ) {
                                        if (infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")) {
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
                            }
                        }
                    }
                }
                if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getTableCells().size() == 4) {
                    if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")) {
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("не определен");
                        boolean checkExceptBact = false;
                        for (int t = 0; t < ExceptionList.exceptBact.size(); t++) {
                            if (ExceptionList.exceptBact.get(t).get(0).equals(workbook.getTables().get(0).getRow(1)
                                    .getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())) {
                                checkExceptBact = true;
                                break;
                            }
                        }
                        if (!checkExceptBact) {
                            ExceptionList.exceptBact.add(new ArrayList<>());
                            ExceptionList.exceptBact.get(ExceptionList.exceptBact.size() - 1).add(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText());
                        }
                        MainController.exceptCheck = true;
                    }
                    if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")) {
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("0.0");
                    }
                    if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")) {
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("отсутствует/крайне низкое/не идентифицирован");
                        /*
                        if(MainController.missingOption){
                            boolean check = false;
                            for(int p = 0;p<infoList.algs.size();p++){
                                if(infoList.algs.get(p).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText())){

                                }
                            }
                            if (!bacterMissing){
                                infoList.missingExpect.add(new ArrayList<>());
                                infoList.missingExpect.get(infoList.missingBacterCount).add(workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(i).getCell(0).getText());
                                infoList.missingBacterCount++;
                            }else{
                                bacterMissing = false;
                            }
                        }
                        */
                    }
                }
                checkMediumRange = true;
                DecimalFormat df = new DecimalFormat("###.##");
                currentStatus += lowLoadStatus;
                double finalCurrentStatus = Double.parseDouble(df.format(currentStatus).replace(",", "."));
                Platform.runLater(new Runnable() {
                    @Override
                    public void run() {
                        mainController.lowLoadText.setText("Дополнительная проверка: " + finalCurrentStatus + "%");
                    }
                });
            }
        }
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                mainController.lowLoadText.setText("");
            }
        });
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
                        if(infoList.uniqBact.get(d).equals(infoList.algs.get(k).get(3)) && !infoList.algs.get(k).get(4).equals(""))
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

    public void setFiveFormat(InfoList infoList, int numberTable, MainController mainController){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        double lowLoadStatus;
        double currentStatus = 0;
        boolean checkBacterValue;
        lowLoadStatus = 100.0 / workbook.getTables().get(numberTable).getRows().size();
        for(int i = 1; i < workbook.getTables().get(numberTable).getRows().size();i++){
            checkBacterValue = true;
            if (workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5) {
                for (int j = 0; j < genusSpecies.size(); j++) {
                    if ((genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !(genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")))
                            || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(genusSpecies.get(j).get(0)
                            .replace("_", " "))
                            || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(genusSpecies.get(j).get(0)
                            .replace(" ", "_"))) && checkBacterValue

                    ) {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText(genusSpecies.get(j).get(1));
                        checkBacterValue = false;
                        for (int k = 0; k < infoList.algs.size(); k++) {
                            if ((workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace("_", " "))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace(" ", "_")))
                            ) {
                                if (workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).getText().equals("0.0")) {
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.algs.get(k).get(2));
                                    break;
                                } else {
                                    if (!infoList.algs.get(k).get(1).equals("0.0")) {
                                        if (checkValueRange(infoList.algs.get(k).get(1), workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).getText())) {
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
                        if (workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5) {
                            if (workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")) {
                                for (int k = 0; k < infoList.algs.size(); k++) {
                                    if (workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                            || infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                            && infoList.algs.get(k).get(0).contains("/")
                                            && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")
                                            || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                            .replace("_", " "))
                                            || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                            .replace(" ", "_"))
                                    ) {
                                        if (infoList.algs.get(k).get(2).equals("среднее значение") || infoList.algs.get(k).get(2).equals("Среднее значение")) {
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
            }
            if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("не определен");
                    boolean checkExceptBact = false;
                    for(int t = 0; t< ExceptionList.exceptBact.size();t++)
                    {
                        if(ExceptionList.exceptBact.get(t).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())){
                            checkExceptBact = true;
                            break;
                        }
                    }
                    if(!checkExceptBact){
                        ExceptionList.exceptBact.add(new ArrayList<>());
                        ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                    }
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
            DecimalFormat df = new DecimalFormat("###.##");
            currentStatus += lowLoadStatus;
            double finalCurrentStatus = Double.parseDouble(df.format(currentStatus).replace(",", "."));
            Platform.runLater(new Runnable() {
                @Override
                public void run() {
                    mainController.lowLoadText.setText("Дополнительная проверка: " + finalCurrentStatus + "%");
                }
            });
        }
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                mainController.lowLoadText.setText("");
            }
        });
    }

    public void setFourFormat(InfoList infoList, int numberTable, MainController mainController){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        double lowLoadStatus;
        double currentStatus = 0;
        lowLoadStatus = 100.0 / workbook.getTables().get(numberTable).getRows().size();
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
                                        && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                        || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                        .replace("_", " "))
                                        || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                        .replace(" ", "_")))
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
                    boolean checkExceptBact = false;
                    for(int t = 0; t< ExceptionList.exceptBact.size();t++)
                    {
                        if(ExceptionList.exceptBact.get(t).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())){
                            checkExceptBact = true;
                            break;
                        }
                    }
                    if(!checkExceptBact){
                        ExceptionList.exceptBact.add(new ArrayList<>());
                        ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                    }
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
            DecimalFormat df = new DecimalFormat("###.##");
            currentStatus += lowLoadStatus;
            double finalCurrentStatus = Double.parseDouble(df.format(currentStatus).replace(",", "."));
            Platform.runLater(new Runnable() {
                @Override
                public void run() {
                    mainController.lowLoadText.setText("Дополнительная проверка: " + finalCurrentStatus + "%");
                }
            });
        }
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                mainController.lowLoadText.setText("");
            }
        });
    }

    public void setThreeDoubleFormat(InfoList infoList, int numberTable, MainController mainController){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        boolean checkFirst, checkSecond;
        double lowLoadStatus;
        double currentStatus = 0;
        lowLoadStatus = 100.0 / workbook.getTables().get(numberTable).getRows().size();
        for(int i = 1; i < workbook.getTables().get(numberTable).getRows().size();i++){
            checkFirst = true;
            checkSecond = true;
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if((genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                        || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                        && genusSpecies.get(j).get(0).contains("/")
                        && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_")
                        )
                        || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(genusSpecies.get(j).get(0)
                        .replace("_", " "))
                        || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(genusSpecies.get(j).get(0)
                        .replace(" ", "_"))) && checkFirst
                        && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("")
                )
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
                            || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(j).get(0)
                            .replace("_", " "))
                            || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(j).get(0)
                            .replace(" ", "_"))
                    ) && checkSecond
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
                    checkFirst_ = true;
                    checkSecond_ = true;
                    for(int k = 0; k<infoList.algs.size();k++) {
                        if(!workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("") && checkFirst_)
                        {
                            if ((workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace("_", " "))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0)
                                    .replace(" ", "_")))
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
                        if(!workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("") && checkSecond_)
                        {
                            if ((workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals(infoList.algs.get(k).get(0))
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText() + "_"))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals(infoList.algs.get(k).get(0)
                                    .replace("_", " "))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals(infoList.algs.get(k).get(0)
                                    .replace(" ", "_")))
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
                        boolean checkExceptBact = false;
                        for(int t = 0; t< ExceptionList.exceptBact.size();t++)
                        {
                            if(ExceptionList.exceptBact.get(t).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())){
                                checkExceptBact = true;
                                break;
                            }
                        }
                        if(!checkExceptBact){
                            ExceptionList.exceptBact.add(new ArrayList<>());
                            ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                        }
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
                            boolean checkExceptBact = false;
                            for(int t = 0; t< ExceptionList.exceptBact.size();t++)
                            {
                                if(ExceptionList.exceptBact.get(t).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())){
                                    checkExceptBact = true;
                                    break;
                                }
                            }
                            if(!checkExceptBact){
                                ExceptionList.exceptBact.add(new ArrayList<>());
                                ExceptionList.exceptBact.get(ExceptionList.exceptBact.size()-1).add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
                            }
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
            DecimalFormat df = new DecimalFormat("###.##");
            currentStatus += lowLoadStatus;
            double finalCurrentStatus = Double.parseDouble(df.format(currentStatus).replace(",", "."));
            Platform.runLater(new Runnable() {
                @Override
                public void run() {
                    mainController.lowLoadText.setText("Дополнительная проверка: " + finalCurrentStatus + "%");
                }
            });
        }
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                mainController.lowLoadText.setText("");
            }
        });
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
                        if(infoList.uniqBact.get(d).equals(infoList.algs.get(k).get(3)) && !infoList.algs.get(k).get(4).equals(""))
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

    public void setTwoFormatWithSer(InfoList infoList, int numberTable, String choiceTable) throws IOException, ClassNotFoundException {
        ArrayList<ArrayList<String>> result = new ArrayList<>();
        if(choiceTable.equals("genus"))
        {
            result.addAll(infoList.genus);
        } else if (choiceTable.equals("species")){
            result.addAll(infoList.species);
        } else if (choiceTable.equals("family")) {
            result.addAll(infoList.family);
        }
        Collections.sort(result, new Comparator<ArrayList<String>>() {
            @Override
            public int compare(ArrayList<String> o1, ArrayList<String> o2) {
                return Double.compare(Double.parseDouble(o2.get(1)), Double.parseDouble(o1.get(1)));
            }
        });
        for(int i = 0; i < result.size();i++){
            workbook.getTables().get(numberTable).createRow();

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1)
                    .setWidth(String.valueOf(workbook.getTables().get(numberTable).getRow(0).getCell(0).getWidth()));
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1)
                    .setWidthType(workbook.getTables().get(numberTable).getRow(0).getCell(1).getWidthType());

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0)
                    .setWidth(String.valueOf(workbook.getTables().get(numberTable).getRow(0).getCell(0).getWidth()));
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1)
                    .setWidthType(workbook.getTables().get(numberTable).getRow(0).getCell(0).getWidthType());

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0)
                    .setIndentationLeft(0);
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0)
                    .setIndentationLeft(200);

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0)
                    .setIndentationRight(0);
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0)
                    .setIndentationRight(0);

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0)
                    .setIndentationFirstLine(0);

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0)
                    .setIndentationFirstLine(0);

            XWPFRun run = workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0).createRun();
            run.removeCarriageReturn();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setItalic(true);
            run.setText(result.get(i).get(0));
            run = workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setText(result.get(i).get(1));
        }
    }

    void loadSortedTable(String nameFileSer) throws IOException, ClassNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(Application.rootDirPath + "\\saveSortedTable_" + nameFileSer + ".ser");
        ObjectInputStream objectInputStream = new ObjectInputStream(fileInputStream);
        sortedTable = (SortedTable) objectInputStream.readObject();
    }

    void saveSortedTable(InfoList infoList, int numberTable, String nameFileSer) throws IOException, ClassNotFoundException {
        sortedTable = new SortedTable();
        for(int i = 0; i < workbook.getTables().get(numberTable).getRows().size();i++) {
            if(!workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("Классификация")
                    && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("РОД БАКТЕРИЙ")
                    && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("ВИД БАКТЕРИЙ")
                    && !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(""))
            {
                sortedTable.tableFirst.add(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText());
            }
        }
        FileOutputStream outputStream = new FileOutputStream(Application.rootDirPath + "\\saveSortedTable_" + nameFileSer + ".ser");
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
        workbookTemp.write(new FileOutputStream(new File(Application.rootDirPath + "\\" + nameObr + ".docx")));
        workbookTemp.close();
    }
}

