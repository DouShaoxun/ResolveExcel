package cn.cruder;

import cn.hutool.core.date.DatePattern;
import cn.hutool.core.date.DateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Date;

/**
 * @author cruder
 * @date 2020-12-08 20:54
 */
public class ResolveAirportDataExcel {

    private final static String EXCEL_DIR = "excel";
    private final static String OUTPUT_DIR = "output";


    public static void main(String[] args) throws IOException, InvalidFormatException {
        String fileName = "AirportData.xlsx";
        String excelPath = getExcelPath(fileName);
        System.out.println(excelPath);
        // 读入excel文件
        File file = new File(excelPath);
        InputStream inputStream = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(inputStream);
        // 数据所在sheet 下标为1
        Sheet sheetAt = workbook.getSheetAt(1);
        Integer beginIndex = 1;
        int endIndex = sheetAt.getLastRowNum();
        StringBuffer start = new StringBuffer("START TRANSACTION ;");

        String foutFileName = DateUtil.format(new Date(), DatePattern.PURE_DATETIME_FORMAT) + ".sql";
        String sqlPath = getOutPutPath(foutFileName);
        System.out.println(sqlPath);

        File fout = new File(sqlPath);
        FileOutputStream fos = new FileOutputStream(fout);

        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
        bw.write(start.toString());
        bw.newLine();


        for (int i = beginIndex; i <= endIndex; i++) {
            //StringBuffer sql = new StringBuffer("UPDATE `u2c_basic`.`config_airport`  SET `gmt` = '0800', `dst` = '0600' WHERE `airport_code` = 'AAB';");
            StringBuffer sql = new StringBuffer("UPDATE `u2c_basic`.`config_airport`  SET ");

            Row row = sheetAt.getRow(i);
            Cell cell0 = row.getCell(i);
            Object airport_code = getCellValue(row.getCell(0));
            Object gmt = getCellValue(row.getCell(17));
            Object dst = getCellValue(row.getCell(18));
            sql.append(" `gmt` =").append(" '").append(gmt).append("', `dst` = ").append("'").append(dst).append("' WHERE `airport_code` = ").append("'").append(airport_code).append("';");
            bw.write(sql.toString());
            bw.newLine();

        }
        StringBuffer commit = new StringBuffer("COMMIT;");
        bw.write(commit.toString());
        bw.newLine();
        bw.close();
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
        String resourcesPath = "src" + File.separator + "main" + File.separator + "resources";
        File directory = new File(resourcesPath);
        String courseFile = directory.getCanonicalPath();
        return courseFile + File.separator + EXCEL_DIR + File.separator + fileName;

    }

    private static String getOutPutPath(String fileName) throws IOException {
        String resourcesPath = "src" + File.separator + "main" + File.separator + "resources";
        File directory = new File(resourcesPath);
        String courseFile = directory.getCanonicalPath();
        return courseFile + File.separator + OUTPUT_DIR + File.separator + fileName;

    }
}
