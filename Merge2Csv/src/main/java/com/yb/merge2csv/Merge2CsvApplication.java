package com.yb.merge2csv;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jmx.JmxAutoConfiguration;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.concurrent.atomic.AtomicInteger;

@SpringBootApplication(exclude = JmxAutoConfiguration.class)
public class Merge2CsvApplication implements CommandLineRunner{

    public static void main(String[] args) {
        SpringApplication.run(Merge2CsvApplication.class, args);
    }


    @Override
    public void run(String... args) throws Exception {

        final ArrayList<String[]> csvData1 = importCsvData("./FILES/csv1.csv", ';');
        final ArrayList<String[]> csvData2 = importCsvData("./FILES/csv2.csv", ';');

        String templatePath = "template.xlsx";
        merge2CsvToXlsx( templatePath,"sheet1", "sheet2", "merged", csvData1, csvData2);

    }


    /*
    @param pCsvFilePath : path to the csv file to import
    @param pCsvSeparator : separator used in the csv file ex : ';' or ','
     */
    public ArrayList<String[]> importCsvData(final String pCsvFilePath, final char pCsvSeparator) throws Exception{
           ArrayList<String[]> lines = new ArrayList<>();
            try (BufferedReader reader = new BufferedReader(new FileReader(pCsvFilePath))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    String[] values = line.split(String.valueOf(pCsvSeparator));
                    System.out.println(Arrays.toString(values));
                    lines.add(values);
                }
            }catch (Exception e){
                throw new Exception("Error while reading csv file : " + e.getMessage());
            }
            return lines;
    }

    protected void merge2CsvToXlsx(
            final String pTemplatePath,
            final String pSheetName1,
            final  String pSheetName2,
            final  String pFileName,
            final ArrayList<String[]> pCsvData1,
            final ArrayList<String[]> pCsvData2
            ) throws IOException, InvalidFormatException {

        final Resource resource = new ClassPathResource(pTemplatePath);
        if (resource.exists()) {
            final XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(resource.getInputStream()));
            final XSSFSheet sheet = workbook.getSheet(pSheetName1);
            final XSSFSheet sheet2 = workbook.getSheet(pSheetName2);
            final AtomicInteger currentRowForSheet1 = new AtomicInteger();
            pCsvData1.forEach(line -> {
                final Row row = sheet.createRow(currentRowForSheet1.getAndIncrement());
                for (int i = 0; i < line.length; i++) {
                    row.createCell(i).setCellValue(line[i]);

                }
            });
            final AtomicInteger  currentRowForSheet2 = new AtomicInteger();
            pCsvData2.forEach(line -> {
                final Row row = sheet2.createRow(currentRowForSheet2.getAndIncrement());
                for (int i = 0; i < line.length; i++) {
                    row.createCell(i).setCellValue(line[i]);
                }
            });
            final String fileName = removeInvalidCharacters(
                    pFileName+ "_" +
                            getTimeStampFromDate(new Date(), "yyyyMMdd_HHmmss")
            );
            final String finalFileName = fileName + ".xlsx";

            final FileOutputStream fileOut = new FileOutputStream("./FILES/"+finalFileName);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            // where is stored the file





            System.out.println("Merge done : " + finalFileName);
            return;
        }
        throw new FileNotFoundException("Template file not found : " + pTemplatePath);


    }
        public  String removeInvalidCharacters(final String pFileName){
            return pFileName.replaceAll("[\\\\/:*?\"<>|]", "_");
        }
    public  String getTimeStampFromDate(final Date pDate, final String pDateFormat) {
        return new SimpleDateFormat(pDateFormat).format(pDate);
    }
}
