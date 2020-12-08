package cn.cruder;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author cruder
 * @date 2020-12-08 20:54
 */
public class ResolveExcel {

    private final static String EXCEL_DIR = "excel";


    public static void main(String[] args) throws IOException, InvalidFormatException {
        String fileName = "test.xlsx";
        String excelPath = getExcelPath(fileName);
        System.out.println(excelPath);
        // 读入excel文件
        File file = new File(excelPath);
        InputStream inputStream = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheetAt = workbook.getSheetAt(0);
        Integer beginIndex = 1;
        int endIndex = sheetAt.getLastRowNum();
        StringBuffer sql = new StringBuffer("INSERT INTO `t_a`(`A`, `B`, `C`) VALUES ");
        for (int i = beginIndex; i <= endIndex; i++) {
            Row row = sheetAt.getRow(i);
            Cell cell0 = row.getCell(i);
            Object aValue = getCellValue(row.getCell(0));
            Object bValue = getCellValue(row.getCell(1));
            Object cValue = getCellValue(row.getCell(2));
            sql.append("(");

            sql.append("'").append(aValue).append("'").append(",");

            // 如果是字符
            if (bValue instanceof String) {
                sql.append("'").append(bValue).append("'").append(",");
            } else {
                sql.append(bValue).append(",");
            }
            sql.append("'").append(cValue).append("'").append(")");
            if (i != endIndex) {
                sql.append(",");
            }
        }
        sql.append(";");
        System.out.println(sql);
    }


    private static Object getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        int cellType = cell.getCellType();
        if (CellType.STRING.getCode() == cellType) {
            return cell.getStringCellValue();
        } else if (CellType.NUMERIC.getCode() == cellType) {
            return cell.getNumericCellValue();
        }
        return "";
    }

    /**
     * 获取文件路径
     *
     * @param fileName 文件名
     * @return
     * @throws IOException io异常
     */
    private static String getExcelPath(String fileName) throws IOException {
        // E:\Code\Java\ResolveExcel\src\main\resources\excel\test.xlsx
        String resourcesPath = "src" + File.separator + "main" + File.separator + "resources";
        File directory = new File(resourcesPath);
        String courseFile = directory.getCanonicalPath();
        return courseFile + File.separator + EXCEL_DIR + File.separator + fileName;

    }
}
