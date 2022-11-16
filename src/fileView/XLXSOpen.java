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

    public void getBioIndex(InfoList infoList){
        String checkIndex = workbook.getSheet("BioIndex").getRow(0).getCell(1).getStringCellValue();
        infoList.bioIndex.add(checkIndex);
        infoList.bioIndex.add(workbook.getSheet("BioIndex").getRow(0).getCell(2).getStringCellValue());

    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void getPhylum(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheet("Phylum").getPhysicalNumberOfRows();i++)
        {
            infoList.phylum.add(new ArrayList<>());
            infoList.phylum.get(i).add(workbook.getSheet("Phylum").getRow(i).getCell(0).getStringCellValue());
            String num = workbook.getSheet("Phylum").getRow(i).getCell(1).getStringCellValue();
            infoList.phylum.get(i).add(num);
        }
    }

    public void getGenus(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheet("Genus").getPhysicalNumberOfRows();i++)
        {
            infoList.genus.add(new ArrayList<>());
            infoList.genus.get(i).add(workbook.getSheet("Genus").getRow(i).getCell(0).getStringCellValue());
            String num = workbook.getSheet("Genus").getRow(i).getCell(1).getStringCellValue();
            infoList.genus.get(i).add(num);
        }
    }

    public void getSpecies(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheet("Species").getPhysicalNumberOfRows();i++)
        {
            infoList.species.add(new ArrayList<>());
            infoList.species.get(i).add(workbook.getSheet("Species").getRow(i).getCell(0).getStringCellValue());
            String num = workbook.getSheet("Species").getRow(i).getCell(1).getStringCellValue();
            infoList.species.get(i).add(num);
        }
    }

    public void getFamily(InfoList infoList) throws IOException {
        for(int i = 0; i < workbook.getSheet("Family").getPhysicalNumberOfRows();i++)
        {
            infoList.family.add(new ArrayList<>());
            infoList.family.get(i).add(workbook.getSheet("Family").getRow(i).getCell(0).getStringCellValue());
            String num = workbook.getSheet("Family").getRow(i).getCell(1).getStringCellValue();
            infoList.family.get(i).add(num);
        }
    }

    public void getFileName(InfoList infoList){
        infoList.fileName = fileName;
    }
}