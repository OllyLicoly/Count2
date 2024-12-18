package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Hello world!
 *
 */
public class App 
{
    ArrayList<String> stringArrayList;
    public static void main( String[] args )
    {
        String filePath = "D" +
                ":\\Java\\My\\Count1\\test.xlsx";  // Укажите путь к вашему файлу Excel

        try {
            FileInputStream file = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(file);  // Для файлов формата .xlsx
            Sheet sheet = workbook.getSheetAt(0);  // Открываем первый лист (если нужно другой, укажите индекс)

            // Массив для хранения суммы по каждому столбцу
            double[] columnSums = new double[sheet.getRow(2).getPhysicalNumberOfCells()];

            // Проходим по всем строкам
            for (Row row : sheet) {
                // Пропускаем первую строку, если она заголовочная
                for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                    Cell cell = row.getCell(i);
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        columnSums[i] += cell.getNumericCellValue();
                    }
                }
            }

            // Выводим итоговые суммы по каждому столбцу
            for (int i = 0; i < columnSums.length; i++) {
                System.out.println("Итог для столбца " + (i + 1) + ": " + columnSums[i]);
            }

            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
