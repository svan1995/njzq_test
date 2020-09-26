import org.apache.commons.collections4.list.FixedSizeList;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.*;

public class CompareDemo {
    public static void main(String[] args) throws IOException {
        Map<String, Map<String, Double>> excelData = new HashMap<>();
        Map<String, Map<String, Double>> sqlData = new HashMap<>();
        Workbook workbook = null;
        FileInputStream inputStream = null;
        String model = "C_M02";
        Writer writer = null;
        inputStream = new FileInputStream("C:/Users/CSF/Desktop/南京证券/一般企业_v2.0/compare/"+ model + ".xlsx");
        writer = new FileWriter("C:/Users/CSF/Desktop/南京证券/一般企业_v2.0/compare/"+ model + ".txt");
        workbook = new XSSFWorkbook(inputStream);
        getExcelData(workbook, excelData);
        getSqlData(workbook, sqlData);

        compareData(excelData, sqlData, writer);
        writer.close();
        workbook.close();
    }

    private static void compareData(Map<String, Map<String, Double>> excelData, Map<String, Map<String, Double>> sqlData, Writer writer) throws IOException {
        writer.write("公司名称,指标名称,sql计算结果,excel结果,是否一致\n");
        Set<String> set = new HashSet<>();
        Set<String> set2 = new HashSet<>();
        Set<String> sqlCompanySet = sqlData.keySet();
        for(String name : sqlCompanySet){
            if (excelData.containsKey(name)){
                set2.add(name);
                Map<String, Double> data1 = sqlData.get(name);
                Map<String, Double> data2 = excelData.get(name);
                Set<String> quanSet = data1.keySet();
                for (String quan : quanSet){
                    Double value1 = data1.get(quan);
                    Double value2 = data2.get(quan);
//                    if (!isEqual(value1, value2)){
//                        writer.write(name + "," + quan + "," + String.format("%.4f", value1) + "," + String.format("%.4f", value2) + "\n");
//                        set.add(name);
//                    }
                    boolean flag = isEqual(value1, value2);
                    writer.write(name + "," + quan + "," + value1 + ","
                            + value2 + "," + flag +"\n");

                }
            }
        }
//        writer.write("共有"+set.size()+"家公司存在差异");
//        System.out.println("共有"+set.size()+"家公司存在差异");
        System.out.println("共对比"+set2.size()+"家公司");

    }

    private static boolean isEqual(Double value1, Double value2) {
        if (value1 == null && value2 == null){
            return true;
        }else if ((value1 == null && value2 != null) || (value1 != null && value2 == null)){
            return false;
        }else {
            if (Math.abs((value1-value2)/value2) < 0.01){
                return true;
            }
            return false;
        }

    }

    private static void getExcelData(Workbook workbook, Map<String, Map<String, Double>> excelData) {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        List<String> names = new ArrayList<>();
        for (int i=1; i<row.getPhysicalNumberOfCells();i++){
            names.add(row.getCell(i).getStringCellValue());
        }

        for (int i=1; i<sheet.getPhysicalNumberOfRows();i++){
            row = sheet.getRow(i);
            String name = row.getCell(0).getStringCellValue();
            Map<String, Double> data = new HashMap<>();
            for (int j=1; j<(names.size()+1);j++){
                if (row.getCell(j)==null || row.getCell(j).getCellType() == CellType.STRING){
                    data.put(names.get(j-1), null);
                }else {
                    data.put(names.get(j-1), row.getCell(j).getNumericCellValue());
                }
            }
            if (data.size() != names.size()){
                throw new RuntimeException("错误");
            }
            excelData.put(name, data);

        }

        System.out.println("excel.size = " + excelData.size());
//       System.out.println(excelData.get("中银保险有限公司"));
    }

    private static void getSqlData(Workbook workbook, Map<String, Map<String, Double>> sqlData) {
        Sheet sheet = workbook.getSheetAt(1);
        Row row = sheet.getRow(0);
        List<String> names = new ArrayList<>();
        for (int i=1; i<row.getPhysicalNumberOfCells();i++){
            names.add(row.getCell(i).getStringCellValue());
        }

        for (int i=1; i<sheet.getPhysicalNumberOfRows();i++){
            row = sheet.getRow(i);
            String name = row.getCell(0).getStringCellValue();
            Map<String, Double> data = new HashMap<>();
            for (int j=1; j<(names.size()+1);j++){
                if (row.getCell(j) == null || row.getCell(j).getCellType() == CellType.STRING){
                    data.put(names.get(j-1), null);
                }else {
                    data.put(names.get(j-1), row.getCell(j).getNumericCellValue());
                }
            }
            if (data.size() != names.size()){
                throw new RuntimeException("错误");
            }
            sqlData.put(name, data);
        }
        System.out.println("sql.size = " + sqlData.size());
//        System.out.println(sqlData.get("攀枝花市商业银行股份有限公司"));
    }
}
