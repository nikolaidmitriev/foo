import java.io.*;
import java.util.*;
import java.util.stream.*;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import org.apache.commons.math3.analysis.interpolation.*;
import org.apache.commons.math3.util.*;

public class App {
   private final static String path = "C://Users/admin/Documents/prod/cvc";
   private final static double[] refValues = new double[]{0.01, 0.05, 0.1, 0.5, 1, 3, 5};

   public static void main(String[] args) throws Exception {
      XSSFWorkbook book = new XSSFWorkbook(); 

      XSSFSheet sheet = book.createSheet("Sheet1");

      XSSFRow row;
      int rowIdx = 0;

      for(String name: getFileNames(path)) {
         Transformer trans = new Transformer(path, name);
         LoessInterpolator loess = new LoessInterpolator();
         trans.setUval(loess.smooth(trans.getIval(), trans.getUval()));

         row = sheet.createRow(rowIdx++);

         int refIdx = 0;
         for (int i = 0; i < trans.getIval().length; i++) {
            if (trans.getIval()[i] > refValues[refIdx]) {
               Cell cell = row.createCell(refIdx);
               cell.setCellValue(trans.getU(i, refValues[refIdx++]));
            }
         }
         Cell cell = row.createCell(refIdx);
         cell.setCellValue(trans.getU(trans.getIval().length - 1, refValues[refIdx++]));
         cell = row.createCell(refIdx);
         cell.setCellValue(name.substring(0, name.length() - 5));
      }

      FileOutputStream out = new FileOutputStream(new File("result.xlsx"));
      book.write(out);
      out.close();
   }
   
   private static List<String> getFileNames(String dir) {
      return Stream.of(new File(dir).listFiles())
      .filter(file -> !file.isDirectory())
      .map(File::getName)
      .collect(Collectors.toList());
   }
}
class Transformer {
   private String name;
   private double[] ival;
   private double[] uval;

   public Transformer(String path, String fileName) throws Exception {
      name = fileName;
      System.out.println(fileName);
      FileInputStream fileInput = new FileInputStream(new File(path + "\\" + name));
      XSSFSheet sheet =  new XSSFWorkbook(fileInput).getSheetAt(0);

      ival = new double[sheet.getPhysicalNumberOfRows() - 2];
      uval = new double[sheet.getPhysicalNumberOfRows() - 2];
      
      for (int i = 0; i < ival.length; i++) {
         Row row = (XSSFRow) sheet.getRow(i + 2);
         Cell cell = row.getCell(1);
         uval[i] = cell.getNumericCellValue();
         cell = row.getCell(2);
         ival[i] = cell.getNumericCellValue();
      }

      fileInput.close();
   }
   
   public void setUval(double[] voltage) {
      uval = voltage;
   }

   public double[] getIval() {
      return this.ival;
   }

   public double[] getUval() {
      return this.uval;
   }

   public double getU(int idx, double i) {
      double u;
      if (idx > 0) {
         u = (uval[idx] - uval[idx - 1]) * (i - ival[idx - 1]) / (ival[idx] - ival[idx - 1]) + uval[idx - 1];
      } else {
         u = (uval[idx + 1] - uval[idx]) * (i - ival[idx]) / (ival[idx + 1] - ival[idx]) + uval[idx];
      }

      return Precision.round(u, 3);
   }
}
