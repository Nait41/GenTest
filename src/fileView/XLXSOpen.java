package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class XLXSOpen {
    String fileName;
    Workbook workbook;
    public XLXSOpen(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        fileName = file.getName();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void getPhylum(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
        {
            infoList.phylum.add(new ArrayList<>());
            infoList.phylum.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
            Double num = workbook.getSheetAt(0).getRow(i).getCell(1).getNumericCellValue();
            infoList.phylum.get(i).add(num.toString());
        }
    }

    public void getGenus(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheetAt(3).getPhysicalNumberOfRows();i++)
        {
            infoList.genus.add(new ArrayList<>());
            infoList.genus.get(i).add(workbook.getSheetAt(3).getRow(i).getCell(0).getStringCellValue());
            Double num = workbook.getSheetAt(3).getRow(i).getCell(1).getNumericCellValue();
            infoList.genus.get(i).add(num.toString());
        }
    }

    public void getSpecies(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheetAt(5).getPhysicalNumberOfRows();i++)
        {
            infoList.species.add(new ArrayList<>());
            infoList.species.get(i).add(workbook.getSheetAt(5).getRow(i).getCell(0).getStringCellValue());
            Double num = workbook.getSheetAt(5).getRow(i).getCell(1).getNumericCellValue();
            infoList.species.get(i).add(num.toString());
        }
    }

    public void getFileName(InfoList infoList){
        infoList.fileName = fileName;
    }
}