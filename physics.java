import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class physics {

    public static void main(String[] args) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sheet");
        int rowid = 0;
        int cellid = 0;
        XSSFRow row0 = sheet.createRow(rowid);
        String[] input_data = {"№", "ti, с", "ti-<t>,c", "(ti-<t>),c^2",
                "Границы интервалов", "dN", "dN/(N*dt), c^-1", "t, c", "ρ, c^-1",
                "<t>±σ", "<t>±2σ", "<t>±3σ",
                "σ", "dTсл", "dTпр", "dT"
        };


        double[] default_numbers = new double[200];
        double average_value = 0;           //сумма значений default_numbers
        for (int i = 0; i < 200; i++) {
            double result = Math.random() + 4.5;
            String str = String.format("%.2f", result);
            str = str.replace(",", ".");
            double result_number = Double.parseDouble(str);
            average_value += result_number;
            default_numbers[i] = result_number;
        }
        average_value = average_value / 200; //среднее значение, но не 2 знака после запятой
        String average_str = String.format("%.2f", average_value);
        average_str = average_str.replace(",", ".");
        double average_number = Double.parseDouble(average_str); // итоговое среднее значение

        double[] third_column = new double[200];
        double[] forth_column = new double[200];
        double value_forth_column = 0;
        double value_third_column = 0;
        for(int i = 0; i < 200; i++){
            third_column[i] = default_numbers[i] - average_number;
            value_third_column += value_third_column;
        }
        for(int i = 0; i < 200; i++){
            forth_column[i] = Math.pow((default_numbers[i] - average_number), 2);
            value_forth_column += forth_column[i];
        }
        double sigmaN = Math.sqrt(value_forth_column/199);



        for (int i = 0; i < 201; i++) {

            XSSFRow row = sheet.createRow(i);
            double min = default_numbers[0];
            for (int z = 1; z < default_numbers.length; z++) {
                if (default_numbers[z] < min) {
                    min = default_numbers[z];
                }
            }
            double max = default_numbers[0];
            for (int z = 1; z < default_numbers.length; z++) {
                if (default_numbers[z] > max) {
                    max = default_numbers[z];
                }
            }
            double default_d = (max - min)/14;


            if (i == 0) {
                for (int j = 0; j < 4; j++) {
                    XSSFCell cell = row.createCell(j);
                    cell.setCellValue(input_data[j]);
                }
                {
                    XSSFCell cell = row.createCell(6);
                    cell.setCellValue("построение гистограммы");
                }
                for (int j = 0; j < 5; j++) { //12 c нуля
                    XSSFCell cell = row.createCell(12 + j);
                    cell.setCellValue(input_data[4 + j]);
                }
            }
            if (i > 0) {
                    XSSFCell _cell = row.createCell(0);
                    _cell.setCellValue(i);
                    _cell = row.createCell(1);
                    _cell.setCellValue(default_numbers[i-1]);
                    _cell = row.createCell(2);
                    _cell.setCellValue(third_column[i-1]);
                    _cell = row.createCell(3);
                    _cell.setCellValue(forth_column[i-1]);


                if (i == 1) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("Min");
                    cell = row.createCell(6);
                    cell.setCellValue(min);
                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min));

                }
                if (i == 2) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("Max");
                    cell = row.createCell(6);
                    cell.setCellValue(max);
                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));
                    cell = row.createCell(13);
                    cell.setCellValue(count(min, min + default_d * (int) (i / 2), default_numbers));
                    cell = row.createCell(14);
                    cell.setCellValue(count(min, min + default_d * (int) (i / 2), default_numbers) / (200 * default_d));
                    cell = row.createCell(15);

                    cell.setCellValue((2*min + default_d * (int) (i / 2) + default_d*((int)((i-1)/2)) )/ 2);
                    cell = row.createCell(16);
                    cell.setCellValue((1/(Math.sqrt(2*Math.PI) * sigmaN)) * Math.exp(-1*Math.pow((5 - for_pow(min, default_d,i)), 2) / (2 * sigmaN * sigmaN)));

                }
                if (i == 3) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("N");
                    cell = row.createCell(6);
                    cell.setCellValue(14);
                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));
                }
                if (i == 4) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("d");
                    cell = row.createCell(6);
                    cell.setCellValue(default_d);
                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));
                    cell = row.createCell(13);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers));
                    cell = row.createCell(14);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers) / (200 * default_d));
                    cell = row.createCell(15);
                    cell.setCellValue((2*min + default_d * (int) (i / 2) + default_d*((int)((i-1)/2)) )/ 2);
                    cell = row.createCell(16);
                    cell.setCellValue(1 / (Math.sqrt(2*Math.PI) * sigmaN) * Math.exp(-Math.pow(5 - (for_pow(min, default_d,i)), 2) / (2 * sigmaN * sigmaN)));

                }
                if(i > 4 && i%2 == 1 && i <= 25 ) {
                    XSSFCell cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));
                }
                if(i > 4 && i%2 == 0 && i <= 25) {

                    XSSFCell cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));
                    cell = row.createCell(13);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers));
                    cell = row.createCell(14);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers) / (200 * default_d));
                    cell = row.createCell(15);
                    cell.setCellValue((2*min + default_d * (int) (i / 2) + default_d*((int)((i-1)/2)) )/ 2);
                    cell = row.createCell(16);
                    cell.setCellValue(1 / (Math.sqrt(2*Math.PI) * sigmaN) * Math.exp(-Math.pow(5 - for_pow(min, default_d,i), 2) / (2 * sigmaN * sigmaN)));
                }
                if (i == 26) {
                    XSSFCell cell = row.createCell(6);
                    cell.setCellValue("Интервал");
                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));

                    cell = row.createCell(13);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers));
                    cell = row.createCell(14);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers) / (200 * default_d));
                    cell = row.createCell(15);
                    cell.setCellValue((2*min + default_d * (int) (i / 2) + default_d*((int)((i-1)/2)) )/ 2);
                    cell = row.createCell(16);
                    cell.setCellValue(1 / (Math.sqrt(2*Math.PI) * sigmaN) * Math.exp(-Math.pow((5 - for_pow(min, default_d,i)), 2) / (2 * sigmaN * sigmaN)));
                }
                if(i == 27) {
                    XSSFCell cell = row.createCell(6);
                    cell.setCellValue("от");
                    cell = row.createCell(7);
                    cell.setCellValue("до");
                    cell = row.createCell(8);
                    cell.setCellValue("dN");
                    cell = row.createCell(9);
                    cell.setCellValue("dN/N");
                    cell = row.createCell(10);
                    cell.setCellValue("P");
                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) ((i-1) / 2)));
                }
                if( i == 28) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("<t>±σ");
                    cell = row.createCell(6);
                    cell.setCellValue(5 - sigmaN);
                    cell = row.createCell(7);
                    cell.setCellValue(5 + sigmaN);
                    double temp = 0;
                    cell = row.createCell(8);
                    for (int z = 0; z < 200; z++) {
                        if (default_numbers[z] > 5 - sigmaN && default_numbers[z] < 5 + sigmaN) {
                            temp++;
                        }
                    }
                    cell.setCellValue(temp);
                    cell = row.createCell(9);
                    cell.setCellValue(temp / 200);
                    cell = row.createCell(10);
                    cell.setCellValue(0.683);

                    cell = row.createCell(12);
                    cell.setCellValue(my_cast(min + default_d * (int) (i / 2)));
                    cell = row.createCell(13);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers));
                    cell = row.createCell(14);
                    cell.setCellValue(count(min + default_d * (int) ((i-1)/ 2), min + default_d * (int) (i / 2), default_numbers) / (200 * default_d));
                    cell = row.createCell(15);
                    cell.setCellValue((2*min + default_d * (int) (i / 2) + default_d*((int)((i-1)/2)) )/ 2);
                    cell = row.createCell(16);
                    cell.setCellValue(1 / (Math.sqrt(2*Math.PI) * sigmaN) * Math.exp(-Math.pow((5 - ((2*min + default_d * (int) (i / 2) + default_d*((int)((i-1)/2)) )/ 2)), 2) / (2 * sigmaN * sigmaN)));
                }
                if(i == 29) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("<t>±2σ");
                    cell = row.createCell(6);
                    cell.setCellValue(5 - 2 * sigmaN);
                    cell = row.createCell(7);
                    cell.setCellValue(5 + 2 * sigmaN);
                    double temp = 0;
                    cell = row.createCell(8);
                    for (int z = 0; z < 200; z++) {
                        if ((default_numbers[z] >( 5 - 2 * sigmaN)) && (default_numbers[z] <( 5 + 2 * sigmaN))) {
                            temp++;
                        }
                    }
                    cell.setCellValue(temp);
                    cell = row.createCell(9);
                    cell.setCellValue(temp / 200);
                    cell = row.createCell(10);
                    cell.setCellValue(0.954);
                    cell = row.createCell(13);
                    cell.setCellValue("Result");
                }
                if (i == 30) {
                    XSSFCell cell = row.createCell(5);
                    cell.setCellValue("<t>±3σ");
                    cell = row.createCell(6);
                    cell.setCellValue(5 - 3 * sigmaN);
                    cell = row.createCell(7);
                    cell.setCellValue(5 + 3 * sigmaN);
                    double temp = 0;
                    cell = row.createCell(8);
                    for (int z = 0; z < 200; z++) {
                        if ((default_numbers[z] > (5 - 3 * sigmaN)) && (default_numbers[z] < (5 + 3 * sigmaN))) {
                            temp++;
                        }
                    }
                    cell.setCellValue(temp);
                    cell = row.createCell(9);
                    cell.setCellValue(temp / 200);
                    cell = row.createCell(10);
                    cell.setCellValue(0.997);
                    cell = row.createCell(12);
                    cell.setCellValue("σ");
                    cell = row.createCell(13);
                    cell.setCellValue(1.0/(200*199)*value_forth_column);
                }
                if(i == 31){
                    XSSFCell cell = row.createCell(12);
                    cell.setCellValue("dTсл");
                    cell = row.createCell(13);
                    cell.setCellValue((1.0/(200*199)*value_forth_column)*1.984);
                }
                if(i == 32){
                    XSSFCell cell = row.createCell(12);
                    cell.setCellValue("dTпр");
                    cell = row.createCell(13);
                    cell.setCellValue(0.005);
                }
                if(i == 33){

                    XSSFCell cell = row.createCell(12);cell.setCellValue("dT");
                    cell = row.createCell(13);
                    cell.setCellValue(Math.sqrt(0.005*0.005+Math.pow((1.0/(200*199)*value_forth_column)*1.984*1.984, 2)));

                }
                if(i == 34){
                    XSSFCell cell = row.createCell(12);
                    cell.setCellValue("Result");
                    cell = row.createCell(13);
                    cell.setCellValue(average_number +"+-"+Math.sqrt(0.005*0.005+Math.pow(sigmaN*1.984, 2)));
                }
            }
        }
        XSSFRow row = sheet.createRow(201);
        XSSFCell cell201_0 = row.createCell(0);
        cell201_0.setCellValue("Result");
        XSSFCell cell201_1 = row.createCell(1);
        cell201_1.setCellValue(average_number);
        XSSFCell cell201_2 = row.createCell(2);
        cell201_2.setCellValue(value_third_column);
        XSSFCell cell201_3 = row.createCell(3);
        cell201_3.setCellValue(sigmaN);
        row = sheet.createRow(202);
        XSSFCell cell202_3 = row.createCell(3);
        cell202_3.setCellValue(1/(sigmaN*Math.sqrt(Math.PI)));


        FileOutputStream out = new FileOutputStream(new File("physics_lab_101.xlsx"));
        workbook.write(out);
        out.close();

    }

    private static double for_pow(double min, double default_d, int num){
        double number = ((2*min + default_d * (int) (num / 2) + default_d*((int)((num-1)/2)) )/ 2);
        number = my_cast(number);
        return number;
    }

    private static int count(double begin, double end,double[] array){
        int _count = 0;
        for(int i = 0; i < 200; i++){
            if(array[i] >= begin && array[i] <= end){
                _count++;
            }
        }
        return _count;
    }
    private static double my_cast(double num){
        String str_num = String.format("%.2f", num);
        str_num = str_num.replace(",", ".");
        return Double.parseDouble(str_num);
    }
}