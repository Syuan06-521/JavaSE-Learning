import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.util.*;

/**
 * @author wb_sunjiaqi04
 * @version 1.0
 * @description: 导出零部件签样表
 * @date 2023/7/25 21:19
 */
public class ExportExcel {
    public static void main(String[] args) {
        long timeMillis = System.currentTimeMillis();
        File file = ExportExcel.getExcelTemplate("测试部件", timeMillis + "");
        System.out.println(file.getAbsoluteFile());
    }

    /**
     * 复制excel模板文件，写入指定信息并传入临时路径
     * @param partName 部件名称
     * @param partId 部件编号
     * @return file对象
     */
    public static File getExcelTemplate(String partName, String partId) {
        Path path = Paths.get("codebase/ext/mtwit/config/templates/零部件签样表.xlsx");
        File oriFile = new File("codebase/ext/mtwit/config/templates/零部件签样表.xlsx");
        String tempPath = path.toString().substring(0, path.toString().lastIndexOf(File.separator))  + File.separator + "tmp";
        File filePath = new File(tempPath);
        if (!filePath.exists()) {
            filePath.mkdirs();
        }
        File file = new File(filePath + File.separator + partName + "_" + partId + ".xlsx");
        try {
            DateFormat dateFormat = DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.CHINA);
            String date = dateFormat.format(new Date());
            XSSFWorkbook workbook = new XSSFWorkbook(oriFile);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int cellCount = 0;
            XSSFRichTextString text = new XSSFRichTextString("test");
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            XSSFFont font = workbook.createFont();
            font.setFontName("等线");
            font.setBold(false);
            font.setFontHeightInPoints((short) 14);
            cellStyle.setFont(font);
            for (Row cells : sheet) {
                for (Cell cell : cells) {
                    if (cell.getStringCellValue() == null || "".equals(cell.getStringCellValue())) {
                        text.applyFont(font);
                        cell.setCellValue(text);
                    }
                    if (cellCount == 82) {
                        XSSFRichTextString dateStr = new XSSFRichTextString(date);
                        dateStr.applyFont(font);
                        cell.setCellValue(dateStr);
                    }
                    cellCount++;
                }
            }
            FileOutputStream out = null;
            try {
                out = new FileOutputStream(file);
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            }
            try {
                workbook.write(out);
                out.close();
                return file;
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }
    /**
     * 初始化Excel样式
     * @return workbook，sheet等对象的map集合
     */
    public static Map<String, Object> initExcelTemplate() {
        // 定义map
        Map<String, Object> objMap = new HashMap<>();
        // 创建XSSFWorkbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("零部件签样表");
        // 创建文本格式
        XSSFCellStyle textStyle = workbook.createCellStyle();
        XSSFCellStyle textStyleEnd = workbook.createCellStyle();
        textStyle.setBorderBottom(BorderStyle.THIN);
        textStyle.setBorderLeft(BorderStyle.THIN);
        textStyle.setBorderRight(BorderStyle.THIN);
        textStyle.setBorderTop(BorderStyle.THIN);
        textStyleEnd.setBorderLeft(BorderStyle.THIN);
        textStyleEnd.setBorderBottom(BorderStyle.THIN);
        // 水平居中
        textStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 结尾一行设置左对齐
        textStyleEnd.setAlignment(HorizontalAlignment.LEFT);
        textStyleEnd.setVerticalAlignment(VerticalAlignment.CENTER);
        // 自动换行
        textStyle.setWrapText(true);
        textStyleEnd.setWrapText(true);
        // 设置合并单元格
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 8));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 4, 5));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 7, 8));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 8));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 1, 8));
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 1, 8));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 1, 8));
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 1, 8));
        sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 8));
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 1, 8));
        sheet.addMergedRegion(new CellRangeAddress(10, 10, 1, 8));
        sheet.addMergedRegion(new CellRangeAddress(11, 11, 0, 8));
        // 设置第一列列宽
        sheet.setColumnWidth(0, 6500);
        // 全局文字格式
        Font font = workbook.createFont();
        // 首行标题文字格式
        Font fontTitle = workbook.createFont();
        // 尾行文字格式
        Font fontEnd = workbook.createFont();
        font.setFontName("等线");
        fontTitle.setFontName("等线");
        fontEnd.setFontName("等线");
        // 设置加粗
        font.setBold(true);
        fontTitle.setBold(true);
        // 设置文字大小
        font.setFontHeightInPoints((short) 14);
        fontTitle.setFontHeightInPoints((short) 16);
        fontEnd.setFontHeightInPoints((short) 12);
        // 添加模块带map
        objMap.put("workbook", workbook);
        objMap.put("sheet", sheet);
        objMap.put("textStyle", textStyle);
        objMap.put("textStyleEnd", textStyleEnd);
        objMap.put("font", font);
        objMap.put("fontTitle", fontTitle);
        objMap.put("fontEnd", fontEnd);
        return objMap;
    }
    /**
     * 提供模板内容
     * @return 模板内容map集合
     */
    public static Map<String, RichTextString> privideData() {
        Map<String, RichTextString> textMap = new LinkedHashMap<>();
        // excel模板内容,key为excel中的cell坐标字符串
        textMap.put("00", new XSSFRichTextString("零部件签样表"));
        textMap.put("10", new XSSFRichTextString("样品名称"));
        textMap.put("13", new XSSFRichTextString("零部件供应商"));
        textMap.put("16", new XSSFRichTextString("整车厂"));
        textMap.put("20", new XSSFRichTextString("确认项目"));
        textMap.put("30", new XSSFRichTextString("尺寸检验/图纸核对"));
        textMap.put("40", new XSSFRichTextString("安装效果及颜色确认"));
        textMap.put("50", new XSSFRichTextString("内测报告"));
        textMap.put("60", new XSSFRichTextString("外测报告"));
        textMap.put("70", new XSSFRichTextString("签核意见"));
        textMap.put("80", new XSSFRichTextString("签核结论及签名"));
        textMap.put("81", new XSSFRichTextString("☐ 封样                 ☐ 有条件封样                    ☐ 不封样"));
        textMap.put("90", new XSSFRichTextString("日期"));
        textMap.put("100", new XSSFRichTextString("说明：\n" +
                "1、尺寸/外观检验/安装效果/测试报告/材质报告均合格，则评定结果为封样；\n" +
                "2、当样品有影响安全、主要功性能、外测结果（含UV等外观类外测）或其他经评估认为不可备料的不合格项时，判定为不封样，其余可判定为有条件封样；\n" +
                "3、有条件封样的限制条件需明确写明；"));
        return textMap;
    }

    /**
     * 导出零部件签样表
     *
     * @param path 导出文件的存放路径
     * @return File对象
     */
    public File exportExcel(String path) {
        DateFormat dateFormat = DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.CHINA);
        String date = dateFormat.format(new Date());
        // 初始化Excel模板格式
        Map<String, Object> objectMap = ExportExcel.initExcelTemplate();
        // 获取workbook
        XSSFWorkbook workbook = (XSSFWorkbook) objectMap.get("workbook");
        // 获取sheet
        XSSFSheet sheet = (XSSFSheet) objectMap.get("sheet");
        // CellMap
        Map<String, XSSFCell> xssfCells = new HashMap<>();
        // rowList
        ArrayList<XSSFRow> xssfRows = new ArrayList<>();
        for (int i = 1; i < 12; i++) {
            XSSFRow row = sheet.createRow(i);
            xssfRows.add(row);
            for (int j = 0; j < 9; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellStyle(((XSSFCellStyle) objectMap.get("textStyle")));
                xssfCells.put((i - 1) + "" + j, cell);
            }
        }
        // 获取模板内容
        Map<String, RichTextString> textStringMap = ExportExcel.privideData();
        // 加载字体
        Set<String> keySet = textStringMap.keySet();
        for (String key : keySet) {
            // 设置字体
            if (!"100".equals(key) & !"81".equals(key)) {
                textStringMap.get(key).applyFont((Font) objectMap.get("font"));
            }
            // 设置单元格初始内容
            xssfCells.get(key).setCellValue(textStringMap.get(key));
        }
        //设置首尾行相应单元格字体
        textStringMap.get("00").applyFont((Font) objectMap.get("fontTitle"));
        textStringMap.get("100").applyFont((Font) objectMap.get("fontEnd"));
        // 设置待填入字段的默认内容
        xssfCells.get("11").setCellValue("测试样本");
        xssfCells.get("14").setCellValue("测试供应商");
        xssfCells.get("17").setCellValue("测试整车厂");
        xssfCells.get("31").setCellValue("测试");
        xssfCells.get("41").setCellValue("测试");
        xssfCells.get("51").setCellValue("测试");
        xssfCells.get("61").setCellValue("测试");
        xssfCells.get("91").setCellValue(date);
        // 设置行高
        xssfRows.get(0).setHeightInPoints(40);
        xssfRows.get(1).setHeightInPoints(50);
        xssfRows.get(2).setHeightInPoints(30);
        xssfRows.get(3).setHeightInPoints(80);
        xssfRows.get(4).setHeightInPoints(80);
        xssfRows.get(5).setHeightInPoints(80);
        xssfRows.get(6).setHeightInPoints(80);
        xssfRows.get(7).setHeightInPoints(30);
        xssfRows.get(8).setHeightInPoints(120);
        xssfRows.get(9).setHeightInPoints(30);
        xssfRows.get(10).setHeightInPoints(120);
        // 尾行单独设置样式
        xssfCells.get("100").setCellStyle(((XSSFCellStyle) objectMap.get("textStyleEnd")));
        // 导出文件
        File filePath = new File(path);
        if (!filePath.exists()) {
            filePath.mkdirs();
        }
        String fileName = path + File.separator + "零部件签样表.xlsx";
        File file = new File(fileName);
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        try {
            workbook.write(out);
            out.close();
            return file;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
