import data.ExceptionList;
import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class AlgOpen {
    public AlgOpen(InfoList infoList) throws IOException, InvalidFormatException {
        File file = new File("C:\\Program Files\\gentest_obr\\algs.xlsx");
        String filePath = file.getPath();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        for(int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
        {
            infoList.algs.add(new ArrayList<>());
            infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
            if(workbook.getSheetAt(0).getRow(i).getCell(1).getCellType() == CellType.NUMERIC){
                Double num = workbook.getSheetAt(0).getRow(i).getCell(1).getNumericCellValue();
                infoList.algs.get(i).add(num.toString());
            } else {
                infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(1).getStringCellValue().replace(",", "."));
            }
            infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
            if(MainController.genusOption){
                if(workbook.getSheetAt(0).getRow(i).getCell(3) == null && workbook.getSheetAt(0).getRow(i).getCell(0) != null)
                {
                    ExceptionList.genusExceptBact.add(new ArrayList<>());
                    ExceptionList.genusExceptBact.get(ExceptionList.genusExceptBact.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
                    GenusExceptionAnalyzer.genusException = true;
                } else {
                    if (workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue().equals("") && !workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue().equals("")){
                        ExceptionList.genusExceptBact.add(new ArrayList<>());
                        ExceptionList.genusExceptBact.get(ExceptionList.genusExceptBact.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
                        GenusExceptionAnalyzer.genusException = true;
                    }
                }
            }
            if(MainController.descriptionOption){
                if(workbook.getSheetAt(0).getRow(i).getCell(4) == null && workbook.getSheetAt(0).getRow(i).getCell(0) != null)
                {
                    ExceptionList.descriptionExpect.add(new ArrayList<>());
                    ExceptionList.descriptionExpect.get(ExceptionList.descriptionExpect.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
                    ExceptionList.descriptionExpect.get(ExceptionList.descriptionExpect.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
                    DescriptionExceptionAnalyzer.descriptionExcept = true;
                } else {
                    if (workbook.getSheetAt(0).getRow(i).getCell(4).getStringCellValue().equals("") && !workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue().equals("")){
                        ExceptionList.descriptionExpect.add(new ArrayList<>());
                        ExceptionList.descriptionExpect.get(ExceptionList.descriptionExpect.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
                        ExceptionList.descriptionExpect.get(ExceptionList.descriptionExpect.size()-1).add(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
                        DescriptionExceptionAnalyzer.descriptionExcept = true;
                    }
                }
            }
            if (workbook.getSheetAt(0).getRow(i).getPhysicalNumberOfCells() > 3) {
                infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
                if(!infoList.uniqBact.contains(workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue()))
                {
                    infoList.uniqBact.add(workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
                }
            }
            if (workbook.getSheetAt(0).getRow(i).getPhysicalNumberOfCells() > 4) {
                infoList.algs.get(i).add(workbook.getSheetAt(0).getRow(i).getCell(4).getStringCellValue());
            }
        }
        workbook.close();
    }
}
