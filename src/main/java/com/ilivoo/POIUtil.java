package com.ilivoo;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.google.common.collect.Lists;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POITextExtractor;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.LineNumberReader;
import java.io.StringReader;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

public class POIUtil {

    private static String extractText(String filePath) {
        String result;
        try (POITextExtractor extractor = ExtractorFactory.createExtractor(new File(filePath))) {
            result = extractor.getText();
        } catch (Exception e) {
            System.out.println("读取文件失败: " + filePath);
            throw new RuntimeException(e);
        }
        return result;
    }

    public static List<List<String>> readDocKV(String docFile) {
        List<List<String>> result = new ArrayList<>();
        String content = extractText(docFile);
        String[] getValues = System.getProperty("getValues") != null ? System.getProperty("getValues").split(",") : null;
        Path filePath = Paths.get(docFile);
        String fileName = filePath.getFileName().toString().split("\\.")[0];
        String fileDir = filePath.getParent().toString();
        result.add(Lists.newArrayList("文件目录", fileDir, "文件名字", fileName));
        content = content.replaceAll("\\r", "");//去掉\\r，此处注意可以利用\\r完全解析出word中的表格数据
        LineNumberReader reader = new LineNumberReader(new StringReader(content));
        reader.lines().forEach(line -> {
            if (line != null && line.length() > 0) {
                String[] kvs = line.split("\\t");//对于末尾为 \\t　的字符串，无法分得最后
                if (line.endsWith("\t")) {
                    String[] tmpKvs = new String[kvs.length + 1];
                    tmpKvs[kvs.length] = "";
                    System.arraycopy(kvs, 0, tmpKvs, 0, kvs.length);
                    kvs = tmpKvs;
                }
                List<String> keyValue = new ArrayList<>();
                Collections.addAll(keyValue, kvs);
                if (getValues != null && getValues.length > 0) {
                    for (String getValue : getValues) {
                        getValue = StringUtils.strip(getValue);
                        if (Strings.isNullOrEmpty(getValue)) {
                            break;
                        }
                        for (String kv : keyValue) {
                            kv = StringUtils.strip(kv);
                            if (kv.startsWith(getValue)) {
                                result.add(Lists.newArrayList(getValue, kv.substring(getValue.length())));
                            }
                        }
                    }
                }
                if (keyValue.size() > 1 || (keyValue.size() == 1 && StringUtils.strip(keyValue.get(0)).length() > 1)) {
                    result.add(keyValue);
                }
            }
        });
        return result;
    }

    private static Workbook workbook(String filePath) {
        Workbook result;
        try (InputStream is = new FileInputStream(filePath)) {
            FileMagic magic = FileMagic.valueOf(FileMagic.prepareToCheckMagic(new FileInputStream(filePath)));
            switch (magic) {
                case OLE2:
                    result = new HSSFWorkbook(is);
                    break;
                case OOXML:
                    result = new XSSFWorkbook(is);
                    break;
                default:
                    throw new RuntimeException("不能解析文件： " + filePath);
            }
        } catch (Exception e) {
            System.out.println("读取文件失败: " + filePath);
            throw new RuntimeException(e);
        }
        return result;
    }

    public static List<String> readExcelLine(String filePath, int line) {
        List<String> result = new ArrayList<>();
        try (Workbook workbook = workbook(filePath)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(line);
            row.cellIterator().forEachRemaining(cell -> {
                String cellValue = cell.getStringCellValue();
                if (!Strings.isNullOrEmpty(cellValue)) {
                    cellValue = StringUtils.strip(cellValue);
                    result.add(cellValue);
                }
            });
        } catch (Exception e) {
            System.out.println("读取文件失败: " + filePath);
            throw new RuntimeException(e);
        }
        return result;
    }

    public static void writeExcel(String filePath, List<Map<String, String>> linesKV, List<String>... headers) {
        Preconditions.checkArgument(headers.length > 0);
        for (List<String> header : headers) {
            Preconditions.checkArgument(header.size() == headers[0].size(), "excel头两行列不相等");
        }
        List<List<String>> lines = Lists.newArrayList(headers);
        try (Workbook workbook = workbook(filePath)) {
            for (Map<String, String> lineKV : linesKV) {
                List<String> line = new ArrayList<>();
                for (String headColumn : headers[0]) {
                    String cellValue = lineKV.get(headColumn);
                    line.add(cellValue);
                }
                lines.add(line);
            }
            int rowNum = 0;
            Sheet sheet = workbook.createSheet();
            for (List<String> line : lines) {
                Row row = sheet.createRow(rowNum);
                for (int columnNum = 0; columnNum < line.size(); columnNum++) {
                    Cell cell = row.createCell(columnNum, CellType.STRING);
                    cell.setCellValue(line.get(columnNum));
                }
                rowNum++;
            }
            workbook.write(new FileOutputStream(FileUtil.newFile(filePath)));
        } catch (Exception e) {
            System.out.println("读取文件失败: " + filePath);
            throw new RuntimeException(e);
        }
    }
}
