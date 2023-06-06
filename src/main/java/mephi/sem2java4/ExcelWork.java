package mephi.sem2java4;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWork {
    
    private final ArrayList<double[]> samples = new ArrayList<>();
          
    public ArrayList<double[]> readFile(String path) throws IOException 
    {
        ArrayList<Double> sampleX = new ArrayList<>();
        ArrayList<Double> sampleY = new ArrayList<>();
        ArrayList<Double> sampleZ = new ArrayList<>();

        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(path));
        XSSFSheet sheet = workbook.getSheet("Вариант 9");
        Iterator<Row> iterator = sheet.iterator();

        while(iterator.hasNext()) 
        {
            Row row = iterator.next();
            Cell cellX = row.getCell(0);
            Cell cellY = row.getCell(1);
            Cell cellZ = row.getCell(2);

            if (cellX.getCellType() == CellType.NUMERIC) {
                sampleX.add(cellX.getNumericCellValue());
            }

            if (cellY.getCellType() == CellType.NUMERIC) {
                sampleY.add(cellY.getNumericCellValue());
            }

            if (cellZ.getCellType() == CellType.NUMERIC) {
                sampleZ.add(cellZ.getNumericCellValue());
            }
        }

        samples.add(sampleX.stream().mapToDouble(Double::doubleValue).toArray());
        samples.add(sampleY.stream().mapToDouble(Double::doubleValue).toArray());
        samples.add(sampleZ.stream().mapToDouble(Double::doubleValue).toArray());
        
        return samples;
    }

    public void writeToFile(DefaultTableModel table, double[][] cov_table, String path) throws FileNotFoundException, IOException
    {
        XSSFWorkbook workbook1 = new XSSFWorkbook();
        XSSFSheet sheet1 = workbook1.createSheet("Calculations");
        Row headerRow1 = sheet1.createRow(0);
        for(int i=0; i<table.getColumnCount(); i++)
        {
            headerRow1.createCell(i).setCellValue(table.getColumnName(i));
        }
        for(int i=0; i<table.getRowCount(); i++)
        {
            XSSFRow row1 = sheet1.createRow(i+1);
            for(int j=0; j<table.getColumnCount(); j++)
            {
                XSSFCell cell1 = row1.createCell(j);
                cell1.setCellValue(table.getValueAt(i, j).toString());
            }
        }
        
    XSSFSheet sheet2 = workbook1.createSheet("Covariance matrix");
            Row headerRow2 = sheet2.createRow(0);
            headerRow2.createCell(1).setCellValue("X");
            headerRow2.createCell(2).setCellValue("Y");
            headerRow2.createCell(3).setCellValue("Z");
            Row RowX = sheet2.createRow(1);
                Cell cellX = RowX.createCell(0);
                cellX.setCellValue("X");
                Cell cell1 = RowX.createCell(1);
                cell1.setCellValue(cov_table[0][0]);

            Row RowY = sheet2.createRow(2);
                Cell cellY = RowY.createCell(0);
                cellY.setCellValue("Y");
                Cell cell2 = RowY.createCell(2);
                cell2.setCellValue(cov_table[1][1]);
                Cell cell4 = RowY.createCell(1);
                cell4.setCellValue(cov_table[0][1]);                 

            Row RowZ = sheet2.createRow(3);
                Cell cellZ = RowZ.createCell(0);
                cellZ.setCellValue("Z");
                Cell cell3 = RowZ.createCell(3);
                cell3.setCellValue(cov_table[2][2]);
                Cell cell5 = RowZ.createCell(1);
                cell5.setCellValue(cov_table[0][2]);
                Cell cell6 = RowZ.createCell(2);
                cell6.setCellValue(cov_table[1][2]);
       
        workbook1.write(new FileOutputStream(path));
        workbook1.close();
    }
}


