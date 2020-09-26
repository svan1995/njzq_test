import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;

public class Excel2Sql {

    private static DecimalFormat df = new DecimalFormat("#.00000");
    public static void main(String[] args) throws IOException {

        Workbook workbook = null;
        FileInputStream inputStream = null;
        String model = "C_M31";
        Writer writer = null;
        writer = new FileWriter("C:/Users/CSF/Desktop/定量指标sql/"+ model + ".sql");
        inputStream = new FileInputStream("C:/Users/CSF/Desktop/定量指标sql/"+ model + ".xlsx");
        workbook = new XSSFWorkbook(inputStream);

        weight_quan(workbook, model, writer);
        weight_qual(workbook, model, writer);
        map_quan(workbook, model, writer);
        map_qual(workbook, model, writer);
        null_map_quan(workbook, model, writer);
        null_map_qual(workbook, model, writer);
        if (workbook.getNumberOfSheets() > 6){
            total_weight(workbook, model, writer);
        }

        writer.close();
        workbook.close();
        inputStream.close();
    }

    private static void weight_quan(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(0);
        writer.write("delete from ods_model_weight_quan t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_model_weight_quan values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getStringCellValue();
            sb.append("'" + item + "', ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

    private static void weight_qual(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(1);
        writer.write("delete from ods_model_weight_qual t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_model_weight_qual values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getStringCellValue();
            sb.append("'" + item + "', ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

    private static void map_quan(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(2);
        writer.write("delete from ods_quan_gear_map_to_score t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_quan_gear_map_to_score values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getStringCellValue();
            sb.append("'" + item + "', ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + ", ");
            cell = row_0.getCell(2);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + ", ");
            cell = row_0.getCell(3);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

    private static void map_qual(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(3);
        writer.write("delete from ods_qual_gear_map_to_score t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_qual_gear_map_to_score values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getStringCellValue();
            sb.append("'" + item + "', ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.0f",item);
            sb.append(item + ", ");
            sb.append("null, ");
            cell = row_0.getCell(2);
            item = cell.getNumericCellValue();
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

    private static void null_map_quan(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(4);
        writer.write("delete from ods_null_quan_gear_map_to_score t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_null_quan_gear_map_to_score values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getStringCellValue();
            sb.append("'" + item + "', ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

    private static void null_map_qual(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(5);
        writer.write("delete from ods_null_qual_gear_map_to_score t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_null_qual_gear_map_to_score values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getStringCellValue();
            sb.append("'" + item + "', ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

    private static void total_weight(Workbook workbook, String model, Writer writer) throws IOException {
        Sheet sheet_0 = workbook.getSheetAt(6);
        writer.write("delete from ods_model_weight_total t where t.model_expo = '" + model + "';\n");
        for (int i=0; i<sheet_0.getPhysicalNumberOfRows();i++){
            StringBuilder sb = new StringBuilder();
            sb.append("insert into table ods_model_weight_total values('");
            sb.append(model + "', ");
            Row row_0 = sheet_0.getRow(i);
            Cell cell = row_0.getCell(0);
            Object item;
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + ", ");
            cell = row_0.getCell(1);
            item = cell.getNumericCellValue();
            item = String.format("%.5f", item);
            sb.append(item + " );");
            writer.write(sb.toString() + "\n");
        }
        writer.write("\n");
    }

}
