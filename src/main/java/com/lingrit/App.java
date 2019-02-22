package com.lingrit;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import com.lingrit.config.Sheet;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.XmlUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

/**
 * Excel Parser
 * @author fxb
 */
public class App {

  private static String CONFIG_PATH = "config.xml";

  private static String EXCEL_PATH = "excel";

  private static String OUTPUT_PATH = "output.xlsx";

  public static void main(String[] args) {

    CONFIG_PATH = System.getProperty("user.dir").concat(File.separator).concat(CONFIG_PATH);
    EXCEL_PATH = System.getProperty("user.dir").concat(File.separator).concat(EXCEL_PATH);
    OUTPUT_PATH = System.getProperty("user.dir").concat(File.separator).concat(OUTPUT_PATH);

    //读取参数多个excel
    File[] files = FileUtil.ls(EXCEL_PATH);
    //读取解析规则
    List<Sheet> sheetList = readConfigFile(CONFIG_PATH);
    //解析多个excel转list<object>
    List<List<String>> rows = parseByConfigFiles(sheetList, files);
    writeToExcel(OUTPUT_PATH, rows);
  }

  private static void writeToExcel(String filePath, List<List<String>> rows) {
    ExcelWriter writer = ExcelUtil.getWriter(filePath);
    writer.passCurrentRow();
    writer.write(rows, true);
    writer.close();
  }


  private static List<Sheet> readConfigFile(String configPath) {
    Document document = XmlUtil.readXML(configPath);
    NodeList childNodes = document.getElementsByTagName("sheet");
    List<Sheet> xmlConfigList = new ArrayList<>();
    for (int i = 0; i < childNodes.getLength(); i++) {
      Element node = (Element) childNodes.item(i);

      int index = Integer.parseInt(node.getElementsByTagName("index").item(0).getFirstChild().getNodeValue());
      int x = Integer.parseInt(node.getElementsByTagName("x").item(0).getFirstChild().getNodeValue());
      int y = Integer.parseInt(node.getElementsByTagName("y").item(0).getFirstChild().getNodeValue());

      xmlConfigList.add(new Sheet(index, x, y));
    }
    return xmlConfigList;
  }

  private static List<List<String>> parseByConfigFiles(List<Sheet> sheetList, File[] files) {
    List<List<String>> rows = new ArrayList<>();
    Arrays.asList(files).stream().forEach(f -> {
      List<String> row = parseByConfig(sheetList, f);
      rows.add(row);
    });
    return rows;
  }

  public static List<String> parseByConfig(List<Sheet> xmlConfigList, File file) {
    List<String> rows = new ArrayList<>();
    xmlConfigList.forEach(config -> {
      ExcelReader reader = ExcelUtil.getReader(file, config.getIndex());
      Object value = reader.readCellValue(config.getX(), config.getY());
      rows.add((String) value);
    });
    return rows;
  }
}
