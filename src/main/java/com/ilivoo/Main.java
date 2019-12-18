package com.ilivoo;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {

    public static void usage() {
        System.out.println("java -jar data-entry.jar excel docs...");
        System.out.println("使用系统变量: -DprintDoc=true, 打印实际读取的值");
        System.out.println("使用系统变量如: -DgetValues=水库简介, 获取水库简介后面的值");
        System.out.println("列标识: key_index#valueIndex|explain");
        System.out.println("列标识: key 列头的标识，如：水库名称");
        System.out.println("列标识: index　当存在多个相同的标识的时候需要使用下标来区分，默认从0开始");
        System.out.println("列标识: valueIndex 值对应key的位置，默认是行对应(valueIndex从0开始)，也可以从列往下对于（valueIndex从-1开始）");
        System.out.println("列标识: explain 注释");
        System.out.println("头必须存在两列，第一列是key的headers，第二行为用户自己指定，但必须存在");
        System.out.println("headers中添加了两个预定义header，可以直接使用(文件目录、文件名字)");
        System.exit(0);
    }

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            usage();
        }
        String excelFile = args[0];
        String[] docFileArray = new String[args.length - 1];
        System.arraycopy(args, 1, docFileArray, 0, docFileArray.length);

        List<Map<String, String>> readDocKV = new ArrayList<>();

        List<String> header0 = POIUtil.readExcelLine(excelFile, 0);
        List<String> header1 = POIUtil.readExcelLine(excelFile, 1);
        List<String> docFiles = FileUtil.listFiles(docFileArray);

        for (String docFile : docFiles) {
            System.out.println("####### 开始读取 " + docFile + " 文件");
            readDocKV.add(readKV(docFile, header0));
        }
        System.out.println("####### 开始写入 " + excelFile + " 文件");
        POIUtil.writeExcel(excelFile, readDocKV, header0, header1);
        System.out.println("####### 已经完成");
    }

    private static void printDocFile(List<List<String>> lineKeyValue) {
        for (List<String> kv : lineKeyValue) {
            System.out.println(kv);
        }
    }

    private static Map<String, String> readKV(String docFile, List<String> header) {
        Map<String, String> result = new HashMap<>();
        List<List<String>> lineKeyValue = POIUtil.readDocKV(docFile);
        if ("true".equalsIgnoreCase(System.getProperty("printDoc"))) {
            printDocFile(lineKeyValue);
        }
        for (String headColumn : header) {
            String[] keyIndexVIndexComment = headColumn.split("\\|");
            Preconditions.checkArgument(keyIndexVIndexComment.length > 0 && keyIndexVIndexComment.length <= 2);

            String[] keyIndexVIndex = keyIndexVIndexComment[0].split("#");
            Preconditions.checkArgument(keyIndexVIndex.length > 0 && keyIndexVIndex.length <= 2);
            int vIndex = 1;
            if (keyIndexVIndex.length == 2) {
                vIndex = Integer.valueOf(StringUtils.strip(keyIndexVIndex[1]));
            }

            String[] keyIndex = keyIndexVIndex[0].split("_");
            Preconditions.checkArgument(keyIndex.length > 0 && keyIndex.length <= 2);
            int index = 0;
            if (keyIndex.length == 2) {
                index = Integer.parseInt(StringUtils.strip(keyIndex[1]));
            }

            String key = trimKey(keyIndex[0]);
            result.put(headColumn, findValue(key, index, vIndex, lineKeyValue));
        }

        return result;
    }

    private static String findValue(String key, int index, int vIndex, List<List<String>> lineKeyValue) {
        int findIndex = 0;
        for (int k = 0; k < lineKeyValue.size(); k++) {
            List<String> keyLine = lineKeyValue.get(k);
            for (int i = 0; i < keyLine.size() - 1; i++) {//trash last
                String rawKey = trimKey(keyLine.get(i));
                if (key.equals(rawKey)) {
                    if (index == findIndex) {
                        String value = null;
                        if (vIndex > 0) { // right
                            int realIndex = i + vIndex;
                            if (realIndex < keyLine.size()) {
                                value = keyLine.get(realIndex);
                            }
                        } else if (vIndex < 0) { //down
                            int realLine = k + Math.abs(vIndex);
                            if (realLine < lineKeyValue.size()) {
                                List<String> valueLine = lineKeyValue.get(realLine);
                                value = valueLine.get(i);
                            }
                        } else {
                            throw new RuntimeException("error index for " + key);
                        }
                        return trimValue(value);
                    } else {
                        findIndex++;
                    }
                }
            }
        }
        return null;
    }

    private static String trimKey(String key) {
        return StringUtils.strip(key);
    }

    private static String trimValue(String value) {
        if (Strings.isNullOrEmpty(value)) {
            return null;
        }
        value = StringUtils.strip(value);
        if (Strings.isNullOrEmpty(value)) {
            return null;
        }
        return value;
    }
}
