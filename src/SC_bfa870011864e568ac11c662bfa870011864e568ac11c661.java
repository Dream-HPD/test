package aquintos.ee.metric.temp.calc;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Map.Entry;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;

import aquintos.commonresources.ui.utils.DialogUtils;
import aquintos.metric.javasupport.AbstractCalculatableSourceCode;
import aquintos.mm.commoncore.attributedefinitionextension.MAbstractAttributeValue;
import aquintos.mm.commoncore.attributedefinitionextension.MAbstractEnumEntry;
import aquintos.mm.commoncore.attributedefinitionextension.MLightWeightBooleanAttributeValue;
import aquintos.mm.commoncore.attributedefinitionextension.MLightWeightByteValue;
import aquintos.mm.commoncore.attributedefinitionextension.MLightWeightEnumAttributeValue;
import aquintos.mm.commoncore.attributedefinitionextension.MLightWeightIntegerAttributeValue;
import aquintos.mm.commoncore.attributedefinitionextension.MLightWeightShortAttributeValue;
import aquintos.mm.commoncore.attributedefinitionextension.MLightWeightStringAttributeValue;
import aquintos.mm.eea.administration.MBusType;
import aquintos.mm.eea.administration.MCANFDType;
import aquintos.mm.eea.administration.MCANType;
import aquintos.mm.eea.componentslayer.MBusConnector;
import aquintos.mm.eea.componentslayer.MCANECUInterface;
import aquintos.mm.eea.componentslayer.MComponentPackage;
import aquintos.mm.eea.componentslayer.MDetailedElectricElectronic;
import aquintos.mm.eea.componentslayer.MECU;
import aquintos.mm.eea.componentslayer.MECUInterface;
import aquintos.mm.eea.componentslayer.MSignalConnector;
import aquintos.mm.eea.componentslayer.signaltransmission.MCANFrameTransmission;
import aquintos.mm.eea.componentslayer.signaltransmission.MFrameTransmission;
import aquintos.mm.eea.componentslayer.signaltransmission.MSignalTransmission;
import aquintos.mm.eea.connections.MBusSystem;
import aquintos.mm.eea.connections.MSignalConnection;
import aquintos.mm.eea.enumerations.MJ1939PriorityEnum;
import aquintos.mm.eea.signalpoollayer.MAbstractBusCommunication;
import aquintos.mm.eea.signalpoollayer.MAbstractBusRoutingArtefact;
import aquintos.mm.eea.signalpoollayer.MAbstractBusRoutingArtefactOwner;
import aquintos.mm.eea.signalpoollayer.MAbstractChannelCommunicationContentOwner;
import aquintos.mm.eea.signalpoollayer.MAbstractTransmittableSignal;
import aquintos.mm.eea.signalpoollayer.MCANCommunication;
import aquintos.mm.eea.signalpoollayer.MCANFrame;
import aquintos.mm.eea.signalpoollayer.MChannelCommunication;
import aquintos.mm.eea.signalpoollayer.MDynamicPartAlternative;
import aquintos.mm.eea.signalpoollayer.MFrame;
import aquintos.mm.eea.signalpoollayer.MFrameGatewayRoutingEntry;
import aquintos.mm.eea.signalpoollayer.MHasSignalIPDUAssignment;
import aquintos.mm.eea.signalpoollayer.MLayoutPackage;
import aquintos.mm.eea.signalpoollayer.MLayoutPackageArtefactOwner;
import aquintos.mm.eea.signalpoollayer.MMultiplexedIPDU;
import aquintos.mm.eea.signalpoollayer.MPDUFrameAssignment;
import aquintos.mm.eea.signalpoollayer.MSignal;
import aquintos.mm.eea.signalpoollayer.MSignalIPDU;
import aquintos.mm.eea.signalpoollayer.MSignalIPDUAssignment;
import aquintos.mm.eea.signalpoollayer.MSystemSignal;
import aquintos.mm.eea.signalpoollayer.nwmanagement.MNmNodeCommunication;
import aquintos.mm.eea.signalpoollayer.timing.MCyclicTiming;
import aquintos.mm.eea.signalpoollayer.timing.MEventControlledTiming;
import aquintos.mm.metricmm.attributedefinition.MMetricAttributeDefinition;
import aquintos.mm.metricmm.datatypes.MDataTypeUnit;
import aquintos.mm.metricmm.datatypes.computationmethods.MAbstractConversionSpecification;
import aquintos.mm.metricmm.datatypes.computationmethods.MComputationMethod;
import aquintos.mm.metricmm.datatypes.computationmethods.MLinearConversion;
import aquintos.mm.metricmm.datatypes.computationmethods.MLinearVerbalTableConversion;
import aquintos.mm.metricmm.datatypes.computationmethods.MTableValue;
import aquintos.mm.metricmm.datatypes.computationmethods.MVerbalTableConversion;
import aquintos.mm.metricmm.datatypes.implementationdatatypes.MImplementationDataType;
import aquintos.mm.metricmm.datatypes.implementationdatatypes.MImplementationDataTypePointer;
import aquintos.mm.metricmm.datatypes.implementationdatatypes.MImplementationValue;
import aquintos.mm.metricmm.datatypes.valuespecification.MIntegerLiteral;
import aquintos.mm.metricmm.datatypes.valuespecification.MStringLiteral;
import aquintos.mm.metricmm.datatypes.valuespecification.MValueSpecification;
import aquintos.mm.metricmm.util.AttributeDefinitionUtility;
import ru.novosoft.mdf.ext.MDFObject;
import ru.novosoft.mdf.impl.MDFRichText;

@SuppressWarnings("nls")
public class SC_bfa870011864e568ac11c662bfa870011864e568ac11c661 extends AbstractCalculatableSourceCode {

  private static final String IN_INPUT = "input"; //$NON-NLS-1$

  private static final String PATH_OUT = "pathOut"; //$NON-NLS-1$

  private static final String OUT_RESULT = "result"; //$NON-NLS-1$

  private List<CellRangeAddress> cellRangeAddressList = new ArrayList<>();

  @Override
  public Object calculateResult() {
    warning("追加Reserved行");
    JsonObject jsonObject = new JsonObject();
    Object input = getInput(IN_INPUT);
    if (input instanceof MComponentPackage) {
      jsonObject = buildDataByMComponentPackage((MComponentPackage) input);
    } else if (input instanceof MBusSystem) {
      jsonObject = buildDataByMBusSystem((MBusSystem) input);
    }
    exportExcel(getInput(PATH_OUT, String.class), jsonObject);
    setResult(OUT_RESULT, "OK");
    return null;
  }

  private void exportExcel(String inPathOut, JsonObject jsonObject) {
    JsonArray dataJsonArray = jsonObject.getAsJsonArray("data");
    for (int i = 0; i < dataJsonArray.size(); i++) {
      XSSFWorkbook book = new XSSFWorkbook();
      // 设置顶端标题行、内容行样式
      XSSFCellStyle headerCellStyle = setHeaderStyle(book, null);
      XSSFCellStyle bodyCellStyle = setBodyStyle(book, null);
      XSSFCellStyle bodyLeftTopCellStyle = setBodyLeftTopStyle(book);
      XSSFCellStyle bodyFrameCellStyle = setBodyFrameStyle(book);
      XSSFCellStyle lastBorderBottomStyle = setLastBorderBottomStyle(book);

      // 获取excelName，用于生成文件名和sheet页名
      JsonObject asJsonObject = dataJsonArray.get(i).getAsJsonObject();
      String excelName = asJsonObject.get("excelName").getAsString();
      // 设置sheet页
      XSSFSheet sheet = book.createSheet(excelName);

      // 顶端标题行行
      List<String> headerList = new ArrayList<>();
      // 合并单元格的集合
      JsonArray mergeColumnArray = new JsonArray();

      // 数据信息
      JsonArray sheetInfoJsonArray = asJsonObject.getAsJsonArray("sheetInfo");
      // 行
      int rowNum = 0;
      for (int j = 0; j < sheetInfoJsonArray.size(); j++) {
        JsonObject sheetInfoJsonObject = sheetInfoJsonArray.get(j).getAsJsonObject();
        // 创建行
        XSSFRow row = sheet.createRow(rowNum);
        // 设置第一行行高
        if (rowNum == 0) {
          row.setHeightInPoints(120);
        }
        // 初始化列
        int num = 0;
        // ECU列
        boolean ecuCol = false;

        // 合并单元格的列集合
        JsonObject mergeColumnJson = new JsonObject();
        mergeColumnJson.addProperty("abbreviation", sheetInfoJsonObject.get("abbreviation").getAsString());
        mergeColumnJson.addProperty("tableColLength", 1);
        mergeColumnJson.addProperty("row", rowNum);
        mergeColumnJson.addProperty("type", "liner");

        boolean hasTable = false;
        if (j != 0 && "table".equals(sheetInfoJsonObject.get("type").getAsString())) {
          mergeColumnJson.addProperty("type", "table");
          hasTable = true;
        }
        mergeColumnArray.add(mergeColumnJson);

        // 列
        for (Entry<String, JsonElement> entry : sheetInfoJsonObject.entrySet()) {
          // 创建列
          XSSFCell cell = row.createCell(num);
          String jsonElementKey = entry.getKey();

          // 新增9个列,算上physicalRange，一共10个列
          if ("physicalRange".equals(jsonElementKey) || "normal".equals(jsonElementKey) || "resolution".equals(jsonElementKey)) {
            if ("physicalRange".equals(jsonElementKey)) {
              mergeColumnJson.addProperty("col", num);
            }
            for (int l = 0; l < 9; l++) {
              num = num + 1;
              // 创建列
              row.createCell(num);
            }
          }

          if (j == 0) {
            // 顶端标题行行数据
            cell.setCellValue(entry.getValue().getAsString());
            if (num < 11) {
              // 自定义颜色对象
              XSSFColor color = new XSSFColor();
              // 根据你需要的rgb值获取byte数组
              color.setRGB(intToByteArray(getIntFromColor(255, 255, 153)));
              cell.setCellStyle(setRotationStyle(book, color));
            }
            // 设置Signal Name、Signal Description样式
            if ("signalName".equals(jsonElementKey) || "signalDescription".equals(jsonElementKey) || "physicalRange".equals(jsonElementKey) || "normal".equals(jsonElementKey)
                    || "resolution".equals(jsonElementKey)) {
              // 自定义颜色对象
              XSSFColor color = new XSSFColor();
              // 根据你需要的rgb值获取byte数组
              color.setRGB(intToByteArray(getIntFromColor(255, 255, 153)));
              cell.setCellStyle(setHeaderBottomStyle(book, color));
            }
            // 调整ECU列表格样式
            if ("signalDefault".equals(jsonElementKey)) {
              ecuCol = true;
            } else if ("physicalRange".equals(jsonElementKey)) {
              ecuCol = false;
            }
            if (ecuCol) {
              if (!"signalDefault".equals(jsonElementKey)) {
                sheet.setColumnWidth(num, 1000);
              }
              // 自定义颜色对象
              XSSFColor color = new XSSFColor();
              // 根据你需要的rgb值获取byte数组
              color.setRGB(intToByteArray(getIntFromColor(255, 255, 153)));
              cell.setCellStyle(setRotationStyle(book, color));
            }
            // 最后6列
            if ("cycleTimeFast".equals(jsonElementKey) || "nrOfReption".equals(jsonElementKey) || "delayTime".equals(jsonElementKey) || "spn".equals(jsonElementKey)
                    || "signalType".equals(jsonElementKey) || "signalTransmissionCycle".equals(jsonElementKey)) {
              // 自定义颜色对象
              XSSFColor color = new XSSFColor();
              // 根据你需要的rgb值获取byte数组
              color.setRGB(intToByteArray(getIntFromColor(255, 192, 0)));
              cell.setCellStyle(setHeaderStyle(book, color));
            }

            headerList.add(entry.getKey());
            num = num + 1;
          } else {
            // 每行数据
            if (hasTable && "table".equals(jsonElementKey)) {
              // 构建嵌入的table
              JsonArray tableJsonArray = entry.getValue().getAsJsonArray();
              //                warning(tableJsonArray.toString());
              mergeColumnJson.addProperty("tableColLength", tableJsonArray.size() + 2);

              // table Value Description起始列
              int col = mergeColumnJson.get("col").getAsInt();

              // 填入table数据
              int size = 0;
              for (int k = 0; k < tableJsonArray.size(); k++) {
                JsonObject tableJsonObject = tableJsonArray.get(k).getAsJsonObject();
                if (k == 0) {
                  size = tableJsonObject.get("size").getAsInt();
                }
                rowNum = rowNum + 1;
                XSSFRow tableRow = sheet.createRow(rowNum);
                for (int l = 0; l < size; l++) {
                  XSSFCell tableCell = tableRow.createCell(col + 9 - l);
                  tableCell.setCellValue(tableJsonObject.get("b" + l).getAsString());
                  if (k == 0) {
                    tableCell.setCellStyle(headerCellStyle);
                  } else {
                    XSSFCellStyle createCellStyle = book.createCellStyle();
                    createCellStyle.setWrapText(true);//自动换行
                    createCellStyle.setAlignment(HorizontalAlignment.CENTER);//文字居中
                    createCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    createCellStyle.setBorderRight(BorderStyle.THIN);
                    createCellStyle.setBorderLeft(BorderStyle.THIN);
                    XSSFFont headFontData = book.createFont();
                    headFontData.setFontName("Arial");
                    createCellStyle.setFont(headFontData);
                    if (k == tableJsonArray.size() - 1) {
                      createCellStyle.setBorderBottom(BorderStyle.THIN);
                    }
                    tableCell.setCellStyle(createCellStyle);
                  }
                }
                XSSFCell tableCell = tableRow.createCell(col + 10);
                tableCell.setCellValue(tableJsonObject.get("valueDescription").getAsString());

                // 追加valueDescription表格样式
                XSSFCellStyle createCellStyle = book.createCellStyle();
                createCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                createCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                if (k == 0) {
                  XSSFFont headFont = book.createFont();
                  headFont.setBold(true);//字体加粗
                  createCellStyle.setFont(headFont);
                }
                tableCell.setCellStyle(createCellStyle);
              }
              // 创建table最后一空行
              rowNum = rowNum + 1;
              XSSFRow createRow = sheet.createRow(rowNum);
              for (int k = col; k < col + 30; k++) {
                createRow.createCell(k).setCellStyle(lastBorderBottomStyle);
              }
            } else {
              for (int k = 0; k < headerList.size(); k++) {
                String key = headerList.get(k);
                if (key.equals(jsonElementKey)) {
                  JsonElement jsonElementValue = entry.getValue();
                  if (jsonElementValue.isJsonNull()) {
                    cell.setCellValue("");
                    cell.setCellStyle(bodyCellStyle);
                  } else if (jsonElementValue.getAsJsonPrimitive().isString()) {
                    cell.setCellValue(entry.getValue().getAsString());
                    if ("s".equals(entry.getValue().getAsString())) {
                      // 自定义颜色对象
                      XSSFColor color = new XSSFColor();
                      // 根据你需要的rgb值获取byte数组
                      color.setRGB(intToByteArray(getIntFromColor(243, 12, 12)));
                      cell.setCellStyle(setBodyStyle(book, color));
                    } else {
                      cell.setCellStyle(bodyCellStyle);
                    }
                  } else if (jsonElementValue.getAsJsonPrimitive().isNumber()) {
                    cell.setCellValue(entry.getValue().getAsInt());
                    cell.setCellStyle(bodyCellStyle);
                  }
                  num = num + 1;
                  break;
                }
              }
            }
            // 设置Signal Name、Signal Description样式
            if ("signalName".equals(jsonElementKey) || "signalDescription".equals(jsonElementKey)) {
              cell.setCellStyle(bodyLeftTopCellStyle);
            }
          }
        }
        rowNum = rowNum + 1;
      }

      // 设置默认列宽
      sheet.setDefaultColumnWidth(10);
      // 全局不显示边框，隐藏Excel网格线，默认值为true
      sheet.setDisplayGridlines(false);
      // 设置前11列列宽
      for (int j = 0; j < 11; j++) {
        sheet.setColumnWidth(j, 2000);
      }
      // 设置Signal Name、Signal Description列宽及样式
      sheet.setColumnWidth(11, 12000);
      sheet.setColumnWidth(12, 24000);

      // 冻结最上面的一行
      sheet.createFreezePane(0, 1);

      // 合并单元格
      if (mergeColumnArray.size() > 0) {
        //          warning(mergeColumnArray.toString());
        int firstRow = 0;
        int mergeLength = 0;
        // 是否记录合并第一行行数
        boolean recordFirstRow = true;
        for (int k = 0; k < mergeColumnArray.size(); k++) {
          JsonObject mergeColumnJson = mergeColumnArray.get(k).getAsJsonObject();
          int row = mergeColumnJson.get("row").getAsInt();
          int col = mergeColumnJson.get("col").getAsInt();
          String type = mergeColumnJson.get("type").getAsString();

          // 设置Physical Range ~ Physical Resolution列宽
          for (int l = 0; l <= 29; l++) {
            sheet.setColumnWidth(col + l, 800);
          }

          // 获取要合并的表格列高
          int colLength = mergeColumnJson.get("tableColLength").getAsInt();
          String value = mergeColumnJson.get("abbreviation").getAsString();
          if (StringUtils.isNotBlank(value)) {
            // abbreviation有值才判断是否合并
            if (k + 1 < mergeColumnArray.size()) {
              // 获取下一个abbreviation值，判断当前和下一个是否相同，相同考虑合并，不同不合并
              JsonObject nextMergeColumnJson = mergeColumnArray.get(k + 1).getAsJsonObject();
              String nextValue = nextMergeColumnJson.get("abbreviation").getAsString();
              if (!value.equals(nextValue) && (mergeLength + colLength) > 1) {
                // 合并（总长度+当前行长度）
                buildMerge(sheet, firstRow, mergeLength, row, col, colLength);
                mergeLength = 0;
                recordFirstRow = true;
              } else {
                // 不合并
                if (recordFirstRow) {
                  firstRow = row;
                }
                if (value.equals(nextValue)) {
                  mergeLength = mergeLength + colLength;
                  recordFirstRow = false;
                }
              }
            } else {
              if ("table".equals(type)) {
                // 合并(最后一行是Physical Range下有table)
                buildMergeTable(lastBorderBottomStyle, sheet, firstRow, mergeLength, row, col, colLength);
              } else {
                if (!recordFirstRow) {
                  buildMerge(sheet, firstRow, mergeLength, row, col, colLength);
                }
              }
            }
          } else if (StringUtils.isBlank(value) && "table".equals(type)) {
            // 合并(abbreviation是空，并且Physical Range下有table)
            buildMergeTable(lastBorderBottomStyle, sheet, firstRow, mergeLength, row, col, colLength);
          }

          // 合并单元格的9个列
          if ("liner".equals(type)) {
            // Physical Range ~ Physical Resolution
            CellRangeAddress physicalRangeFirst = new CellRangeAddress(row, row, col, col + 9);
            CellRangeAddress normalRangeFirst = new CellRangeAddress(row, row, col + 10, col + 19);
            CellRangeAddress physicalResolutionFirst = new CellRangeAddress(row, row, col + 20, col + 29);
            sheet.addMergedRegion(physicalRangeFirst);
            sheet.addMergedRegion(normalRangeFirst);
            sheet.addMergedRegion(physicalResolutionFirst);
            //            setCellRangeStyle(1, physicalRangeFirst, sheet);
            //            setCellRangeStyle(1, normalRangeFirst, sheet);
            //            setCellRangeStyle(1, physicalResolutionFirst, sheet);
            cellRangeAddressList.add(physicalRangeFirst);
            cellRangeAddressList.add(normalRangeFirst);
            cellRangeAddressList.add(physicalResolutionFirst);
          } else if ("table".equals(type)) {
            // 删除Physical Range 、 Normal Range 、 Physical Resolution 第一个单元格边框样式
            XSSFRow xssfRow = sheet.getRow(row);
            XSSFCell physicalRangeCol = xssfRow.getCell(col);
            XSSFCell normalRangeCol = xssfRow.getCell(col + 10);
            XSSFCell physicalResolutionCol = xssfRow.getCell(col + 20);
            physicalRangeCol.setCellStyle(bodyFrameCellStyle);
            normalRangeCol.setCellStyle(bodyFrameCellStyle);
            physicalResolutionCol.setCellStyle(bodyFrameCellStyle);

            // 获取table表格列高
            int tableColLength = mergeColumnJson.get("tableColLength").getAsInt();
            // Byte Number ~ Physical Range
            for (int j = 6; j < col; j++) {
              CellRangeAddress cellRange = new CellRangeAddress(row, row + tableColLength - 1, j, j);
              sheet.addMergedRegion(cellRange);
              //              setCellRangeStyle(1, cellRange, sheet);
              cellRangeAddressList.add(cellRange);
            }
            // SPN ~ Signal transmission cycle
            for (int j = col + 33; j < col + 36; j++) {
              CellRangeAddress cellRange = new CellRangeAddress(row, row + tableColLength - 1, j, j);
              sheet.addMergedRegion(cellRange);
              //              setCellRangeStyle(1, cellRange, sheet);
              cellRangeAddressList.add(cellRange);
            }
            // table Value Description合并单元格
            for (int j = row + 1; j < row + tableColLength - 1; j++) {
              CellRangeAddress tableValueDescription = new CellRangeAddress(j, j, col + 10, col + 10 + 16);
              sheet.addMergedRegion(tableValueDescription);
              if (j == row + 1) {
                //                setCellRangeStyle(1, tableValueDescription, sheet);
                cellRangeAddressList.add(tableValueDescription);
              } else if (j == row + tableColLength - 2) {
                setValueDescriptionCellRangeStyle(1, tableValueDescription, sheet, true);
              } else {
                setValueDescriptionCellRangeStyle(1, tableValueDescription, sheet, false);
              }
            }
          }
        }
      }

      // 合并首行Abbreviation和Message Name列
      XSSFRow row = sheet.getRow(0);
      if (row != null) {
        row.getCell(1).setCellValue("Message Name");
        row.getCell(2).setCellValue("");
        CellRangeAddress messageName = new CellRangeAddress(0, 0, 1, 2);
        sheet.addMergedRegion(messageName);
      }

      warning("size= " + cellRangeAddressList.size());
      for (CellRangeAddress cellRangeAddress : cellRangeAddressList) {
        setCellRangeStyle(1, cellRangeAddress, sheet);
      }

      try {
        // 生成excel文件
        String fullpath = inPathOut + File.separator + excelName + "_" + new SimpleDateFormat("yyyyMMdd").format(new Date()) + ".xlsx";
        OutputStream out = new FileOutputStream(fullpath);
        book.write(out);
      } catch (Exception e) {
        throw new RuntimeException("导出写Excel异常！ 错误：" + e.getLocalizedMessage());
      }
    }
    DialogUtils.openInformation("成功", "导出到“" + PATH_OUT + "”路径下的Excel成功！");
  }

  /**
   * 需合并的行，保留第1行的值，将剩余行的值设置为空
   */
  private void removeCellValue(XSSFSheet sheet, int firstRow, int lastRow, int j) {
    for (int i = firstRow + 1; i <= lastRow; i++) {
      XSSFCell cell = sheet.getRow(i).getCell(j);
      if (cell != null) {
        cell.setCellValue("");
      }
    }
  }

  private void buildMerge(XSSFSheet sheet, int firstRow, int mergeLength, int row, int col, int colLength) {
    // Message ID ~ Message Length [Byte]
    for (int j = 0; j < 6; j++) {
      if (0 == mergeLength) {
        CellRangeAddress cellRange = new CellRangeAddress(row, row + colLength - 1, j, j);
        sheet.addMergedRegionUnsafe(cellRange);
        //        setCellRangeStyle(1, cellRange, sheet);
        cellRangeAddressList.add(cellRange);
      } else {
        int lastRow = firstRow + mergeLength + colLength - 1;
        removeCellValue(sheet, firstRow, lastRow, j);
        CellRangeAddress cellRange = new CellRangeAddress(firstRow, lastRow, j, j);
        sheet.addMergedRegionUnsafe(cellRange);
        //        setCellRangeStyle(1, cellRange, sheet);
        cellRangeAddressList.add(cellRange);
      }
    }
    // Cycle Time Fast(ms) 、 Nr.Of Reption 、 Delay Time(ms)
    for (int j = col + 30; j < col + 30 + 3; j++) {
      if (0 == mergeLength) {
        CellRangeAddress cellRange = new CellRangeAddress(row, row + colLength - 1, j, j);
        sheet.addMergedRegionUnsafe(cellRange);
        //        setCellRangeStyle(1, cellRange, sheet);
        cellRangeAddressList.add(cellRange);
      } else {
        int lastRow = firstRow + mergeLength + colLength - 1;
        removeCellValue(sheet, firstRow, lastRow, j);
        CellRangeAddress cellRange = new CellRangeAddress(firstRow, lastRow, j, j);
        sheet.addMergedRegionUnsafe(cellRange);
        //        setCellRangeStyle(1, cellRange, sheet);
        cellRangeAddressList.add(cellRange);
      }
    }
  }

  private void buildMergeTable(XSSFCellStyle lastBorderBottomStyle, XSSFSheet sheet, int firstRow, int mergeLength, int row, int col, int colLength) {
    buildMerge(sheet, firstRow, mergeLength, row, col, colLength);

    // 添加最后一行Physical Range ~ Physical Resolution的底边边框线
    int lastRowNum = sheet.getLastRowNum(); // 最后一行行数
    if (0 != lastRowNum) {
      XSSFRow lastRow = sheet.getRow(lastRowNum);
      for (int j = col; j < col + 30; j++) {
        XSSFCell cell = lastRow.getCell(j);
        if (cell != null) {
          cell.setCellStyle(lastBorderBottomStyle);
        } else {
          XSSFCell createCell = lastRow.createCell(j);
          createCell.setCellStyle(lastBorderBottomStyle);
        }
      }
    }
  }

  private XSSFCellStyle setHeaderStyle(XSSFWorkbook book, XSSFColor color) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    createCellStyle.setWrapText(true);//自动换行
    createCellStyle.setAlignment(HorizontalAlignment.CENTER);//文字居中
    createCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    XSSFFont headFont = book.createFont();
    headFont.setBold(true);//字体加粗
    // headFont.setItalic(true);//字体倾斜
    createCellStyle.setFont(headFont);
    createCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    createCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    createCellStyle.setBorderTop(BorderStyle.THIN);
    //设置右边框线条类型
    createCellStyle.setBorderRight(BorderStyle.THIN);
    //设置下边框线条类型
    createCellStyle.setBorderBottom(BorderStyle.THIN);
    //设置左边框线条类型
    createCellStyle.setBorderLeft(BorderStyle.THIN);
    if (color != null) {
      createCellStyle.setFillForegroundColor(color);
    }
    return createCellStyle;
  }

  /**
   * 标题文字底对齐
   */
  private XSSFCellStyle setHeaderBottomStyle(XSSFWorkbook book, XSSFColor color) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    createCellStyle.setWrapText(true);//自动换行
    createCellStyle.setAlignment(HorizontalAlignment.CENTER);//文字居中
    createCellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
    XSSFFont headFont = book.createFont();
    headFont.setBold(true);//字体加粗
    // headFont.setItalic(true);//字体倾斜
    createCellStyle.setFont(headFont);
    createCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    createCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    createCellStyle.setBorderTop(BorderStyle.THIN);
    createCellStyle.setBorderRight(BorderStyle.THIN);
    createCellStyle.setBorderBottom(BorderStyle.THIN);
    createCellStyle.setBorderLeft(BorderStyle.THIN);
    if (color != null) {
      createCellStyle.setFillForegroundColor(color);
    }
    return createCellStyle;
  }

  private XSSFCellStyle setBodyStyle(XSSFWorkbook book, XSSFColor color) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    createCellStyle.setWrapText(true);//自动换行
    createCellStyle.setAlignment(HorizontalAlignment.CENTER);//文字居中
    createCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
    createCellStyle.setBorderTop(BorderStyle.THIN);
    createCellStyle.setBorderRight(BorderStyle.THIN);
    createCellStyle.setBorderBottom(BorderStyle.THIN);
    createCellStyle.setBorderLeft(BorderStyle.THIN);
    XSSFFont headFontData = book.createFont();
    headFontData.setFontName("Arial");
    createCellStyle.setFont(headFontData);
    createCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    createCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    if (color != null) {
      createCellStyle.setFillForegroundColor(color);
    }
    return createCellStyle;
  }

  /**
   * 文字左上对齐
   */
  private XSSFCellStyle setBodyLeftTopStyle(XSSFWorkbook book) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    createCellStyle.setWrapText(true);//自动换行
    createCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
    createCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    createCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    createCellStyle.setBorderTop(BorderStyle.THIN);
    createCellStyle.setBorderRight(BorderStyle.THIN);
    createCellStyle.setBorderBottom(BorderStyle.THIN);
    createCellStyle.setBorderLeft(BorderStyle.THIN);
    XSSFFont headFontData = book.createFont();
    headFontData.setFontName("Arial");
    createCellStyle.setFont(headFontData);
    return createCellStyle;
  }

  /**
   * 字体旋转
   */
  private XSSFCellStyle setRotationStyle(XSSFWorkbook book, XSSFColor color) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    createCellStyle.setWrapText(true);//自动换行
    createCellStyle.setAlignment(HorizontalAlignment.CENTER); //文字居中
    createCellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM); //对齐方式
    XSSFFont headFont = book.createFont();
    headFont.setBold(true);//字体加粗
    createCellStyle.setFont(headFont);
    createCellStyle.setBorderTop(BorderStyle.THIN);
    createCellStyle.setBorderRight(BorderStyle.THIN);
    createCellStyle.setBorderBottom(BorderStyle.THIN);
    createCellStyle.setBorderLeft(BorderStyle.THIN);
    createCellStyle.setRotation((short) 90);
    createCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    createCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    if (color != null) {
      createCellStyle.setFillForegroundColor(color);
    }
    return createCellStyle;
  }

  /**
   * 最后一行加底边边框
   */
  private XSSFCellStyle setLastBorderBottomStyle(XSSFWorkbook book) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    createCellStyle.setBorderBottom(BorderStyle.THIN);
    return createCellStyle;
  }

  /**
   * 删除边框样式
   */
  private XSSFCellStyle setBodyFrameStyle(XSSFWorkbook book) {
    XSSFCellStyle createCellStyle = book.createCellStyle();
    return createCellStyle;
  }

  /**
   * 设置合并单元格的边框
   */
  private static void setCellRangeStyle(int border, CellRangeAddress cellRangeAddress, Sheet sheet) {
    RegionUtil.setBorderBottom(border, cellRangeAddress, sheet);
    RegionUtil.setBorderTop(border, cellRangeAddress, sheet);
    RegionUtil.setBorderLeft(border, cellRangeAddress, sheet);
    RegionUtil.setBorderRight(border, cellRangeAddress, sheet);
  }

  /**
   * 设置合并table Value Description单元格的边框
   */
  private static void setValueDescriptionCellRangeStyle(int border, CellRangeAddress cellRangeAddress, Sheet sheet, boolean borderBottom) {
    RegionUtil.setBorderLeft(border, cellRangeAddress, sheet);
    RegionUtil.setBorderRight(border, cellRangeAddress, sheet);
    if (borderBottom) {
      RegionUtil.setBorderBottom(border, cellRangeAddress, sheet);
    }
  }

  /**
   * int转byte[]
   */
  public static byte[] intToByteArray(int i) {
    byte[] result = new byte[4];
    result[0] = (byte) ((i >> 24) & 0xFF);
    result[1] = (byte) ((i >> 16) & 0xFF);
    result[2] = (byte) ((i >> 8) & 0xFF);
    result[3] = (byte) (i & 0xFF);
    return result;
  }

  /**
   * rgb转int
   */
  private static int getIntFromColor(int Red, int Green, int Blue) {
    Red = (Red << 16) & 0x00FF0000;
    Green = (Green << 8) & 0x0000FF00;
    Blue = Blue & 0x000000FF;
    return 0xFF000000 | Red | Green | Blue;
  }

  private JsonObject buildDataByMComponentPackage(MComponentPackage mComponentPackage) {
    JsonObject jsonAllObject = new JsonObject();
    JsonArray jsonAllAddArray = new JsonArray();
    JsonArray jsonAllArray = new JsonArray();
    jsonAllObject.add("data", jsonAllAddArray);
    Stream<MDFObject> subtree = mComponentPackage.subtree();
    List<MBusSystem> mBusSystemList = subtree.filter(a -> a instanceof MBusSystem).map(a -> (MBusSystem) a).collect(Collectors.toList());
    for (int i = 0; i < mBusSystemList.size(); i++) {
      MBusSystem mBusSystem = mBusSystemList.get(i);
      buildPvData(jsonAllArray, mBusSystem);

      // 追加Reserved行
      addReservedRow(jsonAllAddArray, jsonAllArray);
    }
    warning(jsonAllObject.toString());
    return jsonAllObject;
  }

  private JsonObject buildDataByMBusSystem(MBusSystem mBusSystem) {
    JsonObject jsonAllObject = new JsonObject();
    JsonArray jsonAllAddArray = new JsonArray();
    JsonArray jsonAllArray = new JsonArray();
    jsonAllObject.add("data", jsonAllAddArray);
    buildPvData(jsonAllArray, mBusSystem);

    // 追加Reserved行
    addReservedRow(jsonAllAddArray, jsonAllArray);

    List<String> ecuList = new ArrayList<>();
    ecuList.add("GW");
    ecuList.add("ABS");
    ecuList.add("ESP");

    JsonObject asJsonObject = jsonAllAddArray.get(0).getAsJsonObject();
    JsonArray asJsonArray = asJsonObject.get("sheetInfo").getAsJsonArray();

    for (int i = 0; i < 500; i++) {
      JsonObject initBodyJsonData = initBodyJsonData(ecuList);
      initBodyJsonData.addProperty("abbreviation", "" + i);
      asJsonArray.add(initBodyJsonData);
    }

    warning(jsonAllObject.toString());
    return jsonAllObject;
  }

  private void addReservedRow(JsonArray jsonAllAddArray, JsonArray jsonAllArray) {
    JsonArray jsonAddArray = new JsonArray();
    if (jsonAllArray.size() > 0) {
      for (JsonElement jsonElement : jsonAllArray) {
        boolean firstRow = true; // 是否是同abbreviation的第一行
        List<String> ecuList = new ArrayList<>();
        JsonObject jsonObject = jsonElement.getAsJsonObject();
        JsonArray jsonArray = jsonObject.getAsJsonArray("sheetInfo");
        for (int i = 0; i < jsonArray.size(); i++) {
          JsonObject jsonObj = jsonArray.get(i).getAsJsonObject();
          if (i == 0) {
            jsonAddArray.add(jsonObj);
            // 初始化ecu集合
            boolean ecu = false;
            for (Entry<String, JsonElement> entry : jsonObj.entrySet()) {
              String jsonElementKey = entry.getKey();
              if ("signalDefault".equals(jsonElementKey)) {
                ecu = true;
                continue;
              }
              if (ecu) {
                ecuList.add(jsonElementKey);
              }
              if ("reserved".equals(jsonElementKey)) {
                ecu = false;
                break;
              }
            }
          } else {
            // 开始处理
            String abbreviation = jsonObj.get("abbreviation").getAsString();
            String bitNumber = jsonObj.get("bitNumber").getAsString();
            String startBit = jsonObj.get("startBit").getAsString();
            boolean isMultiplexor = jsonObj.get("isMultiplexor").getAsBoolean();
            if (StringUtils.isNotBlank(abbreviation) && StringUtils.isNotBlank(bitNumber) && StringUtils.isNotBlank(startBit) && !isMultiplexor) {
              int startBit1 = jsonObj.get("startBit").getAsInt();
              if (firstRow && startBit1 != 1) {
                // 第一行缺少，需要  新增行
                addFirstReservedRow(jsonAddArray, ecuList, jsonObj, startBit1);

                jsonAddArray.add(jsonObj);
                firstRow = false;

                if (i + 1 < jsonArray.size()) {
                  JsonObject nextJsonObj = jsonArray.get(i + 1).getAsJsonObject();
                  String nextAbbreviation = nextJsonObj.get("abbreviation").getAsString();
                  int endBit = Integer.parseInt(bitNumber.replace("..", "!").split("!")[0]); // Bit Number 32..1
                  if (StringUtils.isNotBlank(nextAbbreviation) && StringUtils.isNotBlank(nextJsonObj.get("startBit").getAsString())) {
                    if (abbreviation.equals(nextAbbreviation)) {
                      int nextStartBit = nextJsonObj.get("startBit").getAsInt();
                      if (endBit + 1 < nextStartBit) {
                        // 数字不连续，并且下一个开始数比本行结束+1大，需要  新增行
                        addCentreReservedRow(jsonAddArray, ecuList, jsonObj, endBit, nextStartBit);
                      }
                      firstRow = false;
                    } else {
                      addLastReservedRow(jsonAddArray, ecuList, jsonObj, endBit);
                      firstRow = true;
                    }
                  } else {
                    addLastReservedRow(jsonAddArray, ecuList, jsonObj, endBit);
                    firstRow = true;
                  }
                }
              } else {
                jsonAddArray.add(jsonObj);
                // 第一行不缺少
                if (i + 1 < jsonArray.size()) {
                  JsonObject nextJsonObj = jsonArray.get(i + 1).getAsJsonObject();
                  String nextAbbreviation = nextJsonObj.get("abbreviation").getAsString();
                  int endBit = Integer.parseInt(bitNumber.replace("..", "!").split("!")[0]); // Bit Number 32..1
                  if (StringUtils.isNotBlank(nextAbbreviation) && StringUtils.isNotBlank(nextJsonObj.get("startBit").getAsString())) {
                    if (abbreviation.equals(nextAbbreviation)) {
                      int nextStartBit = nextJsonObj.get("startBit").getAsInt();
                      if (endBit + 1 < nextStartBit) {
                        addCentreReservedRow(jsonAddArray, ecuList, jsonObj, endBit, nextStartBit);
                      }
                      firstRow = false;
                    } else {
                      // 相同abbreviation的最后一行却，需要  新增行
                      addLastReservedRow(jsonAddArray, ecuList, jsonObj, endBit);
                      firstRow = true;
                    }
                  } else {
                    addLastReservedRow(jsonAddArray, ecuList, jsonObj, endBit);
                    firstRow = true;
                  }
                }
              }
            } else {
              jsonAddArray.add(jsonObj);
              firstRow = true;
            }
          }
        }
        JsonObject jsonAddObject = new JsonObject();
        jsonAddObject.addProperty("excelName", jsonObject.get("excelName").getAsString());
        jsonAddObject.add("sheetInfo", jsonAddArray);
        jsonAllAddArray.add(jsonAddObject);
      }
    }
  }

  private void addFirstReservedRow(JsonArray jsonAddArray, List<String> ecuList, JsonObject jsonObj, int startBit1) {
    int endBit = 0;
    addCentreReservedRow(jsonAddArray, ecuList, jsonObj, endBit, startBit1);
  }

  private void addCentreReservedRow(JsonArray jsonAddArray, List<String> ecuList, JsonObject jsonObj, int endBit, int nextStartBit) {
    // startBit
    int startBitAdd = endBit + 1;
    // bitNumber
    String bitNumberAdd = (nextStartBit - 1) + ".." + startBitAdd;
    // signalLength
    int signalLength = nextStartBit - startBitAdd;
    // byteNumber
    String byteNumber = getByteNumberValue(endBit, signalLength);
    if (signalLength > 0) {
      // signalDefault
      String hexStr = getSignalDefaultValue(signalLength);

      bulidAddReservedRow(jsonAddArray, ecuList, jsonObj, signalLength, byteNumber, bitNumberAdd, startBitAdd, hexStr);
    }
  }

  private void addLastReservedRow(JsonArray jsonAddArray, List<String> ecuList, JsonObject jsonObj, int endBit) {
    // startBit
    int startBitAdd = endBit + 1;
    // bitNumber
    int messageLength = jsonObj.get("messageLength").getAsInt();
    String bitNumberAdd = (messageLength * 8) + ".." + startBitAdd;
    // signalLength
    int signalLength = messageLength * 8 - startBitAdd + 1;
    // byteNumber
    String byteNumber = getByteNumberValue(endBit, signalLength);
    if (signalLength > 0) {
      // signalDefault
      String hexStr = getSignalDefaultValue(signalLength);

      bulidAddReservedRow(jsonAddArray, ecuList, jsonObj, signalLength, byteNumber, bitNumberAdd, startBitAdd, hexStr);
    }
  }

  private String getSignalDefaultValue(int signalLength) {
    String signalDefaultStr = "";
    for (int j = 0; j < signalLength; j++) {
      signalDefaultStr = signalDefaultStr + "1";
    }
    BigInteger decimal = new BigInteger(signalDefaultStr, 2);
    String hexStr = "0x" + decimal.toString(16).toUpperCase();
    return hexStr;
  }

  private String getByteNumberValue(int startBit, int signalLength) {
    int byteNumber1 = (int) Math.floor(startBit / 8) + 1;
    int byteNumber2 = (int) Math.floor((startBit + signalLength - 1) / 8) + 1;
    String byteNumber = "";
    if (byteNumber1 == byteNumber2) {
      byteNumber = byteNumber1 + "";
    } else {
      byteNumber = byteNumber1 + ".." + byteNumber2;
    }
    return byteNumber;
  }

  private void bulidAddReservedRow(JsonArray jsonAddArray, List<String> ecuList, JsonObject jsonObj, int signalLength, String byteNumber, String bitNumber, int startBit, String signalDefault) {
    JsonObject bodyJson = new JsonObject();
    bodyJson.addProperty("messageId", jsonObj.get("messageId").getAsString());
    bodyJson.addProperty("abbreviation", jsonObj.get("abbreviation").getAsString());
    bodyJson.addProperty("messageName", jsonObj.get("messageName").getAsString());
    bodyJson.addProperty("cyclic", jsonObj.get("cyclic").getAsString());
    bodyJson.addProperty("sendType", jsonObj.get("sendType").getAsString());
    bodyJson.addProperty("messageLength", jsonObj.get("messageLength").getAsString());
    bodyJson.addProperty("multiplexingValue", "");
    bodyJson.addProperty("byteNumber", byteNumber);
    bodyJson.addProperty("bitNumber", bitNumber);
    bodyJson.addProperty("signalLength", signalLength);
    bodyJson.addProperty("startBit", startBit);
    bodyJson.addProperty("signalName", "Reserved");
    bodyJson.addProperty("signalDescription", "reserved for future extensions");
    bodyJson.addProperty("signalDefault", signalDefault);
    for (String ecuName : ecuList) {
      bodyJson.addProperty(ecuName, jsonObj.get(ecuName).getAsString());
    }
    bodyJson.addProperty("reserved", "");
    bodyJson.addProperty("isSelf", "");
    bodyJson.addProperty("new", "");
    bodyJson.addProperty("physicalRange", "");
    bodyJson.addProperty("normal", "");
    bodyJson.addProperty("resolution", "");
    bodyJson.addProperty("cycleTimeFast", jsonObj.get("cycleTimeFast").getAsString());
    bodyJson.addProperty("nrOfReption", jsonObj.get("nrOfReption").getAsString());
    bodyJson.addProperty("delayTime", jsonObj.get("delayTime").getAsString());
    bodyJson.addProperty("spn", 0);
    bodyJson.addProperty("signalType", "");
    bodyJson.addProperty("signalTransmissionCycle", "");
    bodyJson.addProperty("type", "liner");
    bodyJson.addProperty("table", "");
    jsonAddArray.add(bodyJson);
  }

  private void buildPvData(JsonArray jsonAllArray, MBusSystem mBusSystem) {
    JsonObject jsonObject = new JsonObject();
    JsonArray jsonArray = new JsonArray();
    MBusType busType = mBusSystem.getBusType();
    String excelName = "";
    boolean isCanType = false;
    if (busType instanceof MCANFDType) {
      excelName = mBusSystem.getName() + "_CANFD";
    } else if (busType instanceof MCANType) {
      excelName = mBusSystem.getName() + "_CAN";
      isCanType = true;
    }
    jsonObject.addProperty("excelName", excelName);
    // 初始化表格第一行
    JsonObject headerJson = initFrontHeaderJsonData();

    // Bus Connector
    List<String> ecuList = new ArrayList<>();
    List<MBusConnector> busConnectors = mBusSystem.getBusConnectors();
    if (busConnectors != null && !busConnectors.isEmpty()) {
      for (MBusConnector mBusConnector : busConnectors) {
        MDetailedElectricElectronic electronicComposite = mBusConnector.getElectronicComposite();
        if (electronicComposite instanceof MECU) {
          MECU mecu = (MECU) electronicComposite;
          String name = mecu.getName();
          headerJson.addProperty(name, name);
          ecuList.add(name);
        } else {
          logErrorMessage("错误：" + mBusSystem.getName() + "表格缺失ECU表头列！");
        }
      }
    } else {
      logErrorMessage("错误：" + mBusSystem.getName() + "表格缺失ECU表头列！");
    }
    headerJson = initAfterHeaderJsonData(headerJson, mBusSystem.getName());
    boolean transmissionStatus = false;

    // CAN Communication
    Collection<MAbstractBusCommunication> busCommunications = mBusSystem.getBusCommunications();
    if (busCommunications != null && !busCommunications.isEmpty()) {
      for (MAbstractBusCommunication mAbstractBusCommunication : busCommunications) {
        if (mAbstractBusCommunication instanceof MCANCommunication) {
          MCANCommunication mcanCommunication = (MCANCommunication) mAbstractBusCommunication;
          MAbstractBusRoutingArtefact busRouting = mcanCommunication.getBusRouting();
          if (busRouting instanceof MChannelCommunication) {
            MChannelCommunication mChannelCommunication = (MChannelCommunication) busRouting;
            List<MSignalTransmission> signalTransmissionList = mChannelCommunication.getSignalTransmissions();
            if (signalTransmissionList != null && !signalTransmissionList.isEmpty()) {
              jsonArray.add(headerJson);
              for (MSignalTransmission mSignalTransmission : signalTransmissionList) {
                if (mSignalTransmission.isIncludedInActiveVariant() && mSignalTransmission.isPartOfActiveVariant()) {
                  transmissionStatus = true;
                  // Signal
                  MAbstractTransmittableSignal mAbstractTransmittableSignal = mSignalTransmission.getSignalType();
                  if (mAbstractTransmittableSignal instanceof MSignal) {

                    // 初始化表格数据
                    JsonObject bodyJson = initBodyJsonData(ecuList);

                    int nmNodeId = 0;
                    boolean nameDefaultMarker = true;
                    // ECU Interface --> Sending ECU Interface
                    Collection<MECUInterface> sendingECUInterfaceList = mSignalTransmission.getSendingECUInterface();
                    if (sendingECUInterfaceList != null && !sendingECUInterfaceList.isEmpty()) {
                      for (MECUInterface mecuInterface : sendingECUInterfaceList) {
                        // Connector
                        MSignalConnector connector = mecuInterface.getConnector();
                        if (connector instanceof MBusConnector) {
                          MBusConnector mBusConnector = (MBusConnector) connector;
                          MDetailedElectricElectronic electronicComposite = mBusConnector.getElectronicComposite();
                          if (electronicComposite instanceof MECU) {
                            MECU mecu = (MECU) electronicComposite;
                            String name = mecu.getName();
                            if (ecuList.contains(name)) {
                              bodyJson.addProperty(name, "s");
                            }
                          }
                        }
                      }
                    }

                    // ECU Interface --> Receiving ECU Interface
                    Collection<MECUInterface> receivingECUInterfaceList = mSignalTransmission.getReceivingECUInterface();
                    if (receivingECUInterfaceList != null && !receivingECUInterfaceList.isEmpty()) {
                      for (MECUInterface mecuInterface : receivingECUInterfaceList) {
                        MSignalConnector connector = mecuInterface.getConnector();
                        if (connector instanceof MBusConnector) {
                          MBusConnector mBusConnector = (MBusConnector) connector;
                          MDetailedElectricElectronic electronicComposite = mBusConnector.getElectronicComposite();
                          if (electronicComposite instanceof MECU) {
                            MECU mecu = (MECU) electronicComposite;
                            String name = mecu.getName();
                            if (ecuList.contains(name)) {
                              bodyJson.addProperty(name, "r");
                            }
                          }
                        }
                      }
                    }

                    MSignal mSignal = (MSignal) mAbstractTransmittableSignal;
                    String frameName = "";
                    boolean isMultiplexor = false;

                    // CAN Frame
                    // Signal --> signalIPDUAssignments --> signalIPDU -->frameAssignments --> frame
                    List<MSignalIPDUAssignment> signalIPDUAssignments = mSignal.getSignalIPDUAssignments();
                    if (signalIPDUAssignments != null && !signalIPDUAssignments.isEmpty()) {
                      List<MSignalIPDU> signalIPDUs = signalIPDUAssignments.stream().map(x -> x.getSignalIPDU()).filter(x -> x instanceof MSignalIPDU).map(x -> (MSignalIPDU) x).collect(Collectors.toList());
                      for (MSignalIPDU mSignalIPDU : signalIPDUs) {
                        List<MPDUFrameAssignment> frameAssignments = mSignalIPDU.getFrameAssignments();
                        String frameNameAndNmNodeId = buildFrameData(bodyJson, frameAssignments, mBusSystem, isCanType);

                        if (StringUtils.isBlank(frameName) && StringUtils.isNotBlank(frameNameAndNmNodeId)) {
                          frameName = frameNameAndNmNodeId.split("!")[0];
                          nmNodeId = Integer.parseInt(frameNameAndNmNodeId.split("!")[1]);
                          nameDefaultMarker = Boolean.parseBoolean(frameNameAndNmNodeId.split("!")[2]);
                        }
                      }

                      if (StringUtils.isBlank(frameName)) {
                        int ipduNum = 0;
                        for (MSignalIPDU mSignalIPDU : signalIPDUs) {
                          // Multiplexing/Value 查询
                          Collection<MDynamicPartAlternative> dynamicPartAlternatives = mSignalIPDU.getDynamicPartAlternatives();
                          if (!dynamicPartAlternatives.isEmpty()) {
                            ipduNum = ipduNum + 1;

                            for (MDynamicPartAlternative mDynamicPartAlternative : dynamicPartAlternatives) {
                              MMultiplexedIPDU multiplexedIPDU = mDynamicPartAlternative.getDynamicPart().getMultiplexedIPDU();
                              List<MPDUFrameAssignment> frameAssignments2 = multiplexedIPDU.getFrameAssignments();
                              String frameNameAndNmNodeId = buildFrameData(bodyJson, frameAssignments2, mBusSystem, isCanType);

                              if (StringUtils.isNotBlank(frameNameAndNmNodeId)) {
                                frameName = frameNameAndNmNodeId.split("!")[0];
                                nmNodeId = Integer.parseInt(frameNameAndNmNodeId.split("!")[1]);
                                nameDefaultMarker = Boolean.parseBoolean(frameNameAndNmNodeId.split("!")[2]);
                              }
                            }
                          }
                          // Multiplexing/Value 设置
                          if (ipduNum > 1) {
                            bodyJson.addProperty("multiplexingValue", "Multiplexor");
                            isMultiplexor = true;
                          } else if (ipduNum == 1) {
                            bodyJson.addProperty("multiplexingValue", mSignalIPDU.getName().substring(mSignalIPDU.getName().lastIndexOf("_") + 1));
                            isMultiplexor = true;
                          }
                        }
                      }
                    }
                    bodyJson.addProperty("isMultiplexor", isMultiplexor);

                    // General
                    String signalName = mSignal.getName();
                    MDFRichText description = mSignal.getDescription();
                    int signalLength = mSignal.getBitLength();
                    if (signalLength == 0) {
                      logErrorMessage("错误：signalLength长度必填，不可为空或0，错误signalName= " + signalName);
                    }

                    MValueSpecification rawInitValue = mSignal.getRawInitValue();
                    if (rawInitValue instanceof MIntegerLiteral) {
                      MIntegerLiteral mIntegerLiteral = (MIntegerLiteral) rawInitValue;
                      int signalInitial = mIntegerLiteral.getValue();
                      String hexString = Integer.toHexString(signalInitial);
                      if (StringUtils.isNotBlank(hexString)) {
                        bodyJson.addProperty("signalDefault", "0x" + hexString.toUpperCase());
                      }
                    }
                    bodyJson.addProperty("signalName", signalName);
                    bodyJson.addProperty("signalDescription", description == null ? "" : description.getPlainText());
                    bodyJson.addProperty("signalLength", signalLength);

                    // Implementation Data Type
                    MImplementationValue implementationValue = null;
                    MImplementationDataType implementationType = mSignal.getImplementationType();
                    if (implementationType instanceof MImplementationValue) {
                      implementationValue = (MImplementationValue) implementationType;
                    } else if (implementationType instanceof MImplementationDataTypePointer) {
                      MImplementationDataTypePointer mImplementationDataTypePointer = (MImplementationDataTypePointer) implementationType;
                      MImplementationDataType implementationTypePointer = mImplementationDataTypePointer.getImplementationType();
                      if (implementationTypePointer instanceof MImplementationValue) {
                        implementationValue = (MImplementationValue) implementationTypePointer;
                      }
                    }
                    if (implementationValue != null) {
                      MComputationMethod usedComputationMethod = implementationValue.getUsedComputationMethod();
                      if (usedComputationMethod != null) {
                        boolean isLinerType = true;
                        // Internal to Physical Conversion
                        MAbstractConversionSpecification internalToPhysicalConversion = usedComputationMethod.getInternalToPhysicalConversion();
                        if (internalToPhysicalConversion instanceof MLinearVerbalTableConversion) {
                          // Linear Verbal Table Conversion
                          MLinearVerbalTableConversion mLinearVerbalTableConversion = (MLinearVerbalTableConversion) internalToPhysicalConversion;
                          List<MTableValue> tableValues = mLinearVerbalTableConversion.getTableValues();
                          if (tableValues == null || tableValues.isEmpty()) {
                            buildLinearVerbalTableConversion(mLinearVerbalTableConversion, bodyJson);
                          } else {
                            isLinerType = false;
                            buildVerbalTableConversion(tableValues, bodyJson, signalName);
                          }
                        } else if (internalToPhysicalConversion instanceof MVerbalTableConversion) {
                          // Verbal Table Conversion
                          MVerbalTableConversion mVerbalTableConversion = (MVerbalTableConversion) internalToPhysicalConversion;
                          List<MTableValue> tableValues = mVerbalTableConversion.getTableValues();
                          if (tableValues != null && !tableValues.isEmpty()) {
                            isLinerType = false;
                            buildVerbalTableConversion(tableValues, bodyJson, signalName);
                          }
                        } else if (internalToPhysicalConversion instanceof MLinearConversion) {
                          // Liner Conversion
                          MLinearConversion mLinearConversion = (MLinearConversion) internalToPhysicalConversion;
                          buildLinerConversion(mLinearConversion, bodyJson);
                        }

                        // Physical to Internal Conversion
                        if (isLinerType) {
                          MAbstractConversionSpecification physicalToInternalConversion = usedComputationMethod.getPhysicalToInternalConversion();
                          if (physicalToInternalConversion instanceof MLinearConversion) {
                            MLinearConversion mLinearConversion = (MLinearConversion) physicalToInternalConversion;
                            String min = mLinearConversion.getMin();
                            String max = mLinearConversion.getMax();
                            bodyJson.addProperty("physicalRange", min + ".." + max);
                          } else if (physicalToInternalConversion instanceof MLinearVerbalTableConversion) {
                            MLinearVerbalTableConversion mLinearVerbalTableConversion = (MLinearVerbalTableConversion) physicalToInternalConversion;
                            String min = mLinearVerbalTableConversion.getMin();
                            String max = mLinearVerbalTableConversion.getMax();
                            bodyJson.addProperty("physicalRange", min + ".." + max);
                          }

                          // Data Type Unit
                          String displayName = "";
                          MDataTypeUnit unit = usedComputationMethod.getUnit();
                          if (unit != null) {
                            displayName = unit.getDisplayName();
                          }
                          String physicalResolution = bodyJson.get("resolution").getAsString();
                          if (StringUtils.isNotBlank(physicalResolution)) {
                            bodyJson.addProperty("resolution", physicalResolution + " " + displayName);
                          }
                        }
                      }
                    }

                    // J1939 Specific
                    MSystemSignal systemSignal = mSignal.getSystemSignal();
                    if (systemSignal != null) {
                      int spn = systemSignal.getSpn();
                      bodyJson.addProperty("spn", spn);
                    }

                    // Signal-IPDU-Assignment
                    if (signalIPDUAssignments != null && !signalIPDUAssignments.isEmpty()) {
                      for (MSignalIPDUAssignment mSignalIPDUAssignment : signalIPDUAssignments) {
                        boolean isSame = false;
                        // Signal IPDU
                        MHasSignalIPDUAssignment signalIPDU = mSignalIPDUAssignment.getSignalIPDU();
                        if (signalIPDU instanceof MSignalIPDU) {
                          MSignalIPDU mSignalIPDU = (MSignalIPDU) signalIPDU;
                          if (frameName.equals(mSignalIPDU.getName())) {
                            isSame = true;
                          }
                        }
                        if (isSame || isMultiplexor) {
                          int startBit = mSignalIPDUAssignment.getStartPosition();
                          // Start Position（start bit）
                          String byteNumber = getByteNumberValue(startBit, signalLength);
                          String bitNumber = (startBit + signalLength) + ".." + (startBit + 1);
                          bodyJson.addProperty("startBit", startBit + 1);
                          bodyJson.addProperty("byteNumber", byteNumber);
                          bodyJson.addProperty("bitNumber", bitNumber);

                          // Signal IPDU
                          if (signalIPDU instanceof MSignalIPDU) {
                            MSignalIPDU mSignalIPDU = (MSignalIPDU) signalIPDU;
                            Collection<MDynamicPartAlternative> dynamicPartAlternatives = mSignalIPDU.getDynamicPartAlternatives();
                            if (!dynamicPartAlternatives.isEmpty()) {
                              for (MDynamicPartAlternative mDynamicPartAlternative : dynamicPartAlternatives) {
                                MMultiplexedIPDU multiplexedIPDU = mDynamicPartAlternative.getDynamicPart().getMultiplexedIPDU();
                                boolean j1939PduFormatDefaultMarker = multiplexedIPDU.getJ1939PduFormatDefaultMarker();
                                boolean j1939PduSpecificDefaultMarker = multiplexedIPDU.getJ1939PduSpecificDefaultMarker();
                                MJ1939PriorityEnum j1939Priority = multiplexedIPDU.getJ1939Priority();
                                if (j1939Priority != null && !j1939PduFormatDefaultMarker && !j1939PduSpecificDefaultMarker && !nameDefaultMarker) {
                                  int j1939PriorityInt = Integer.valueOf(j1939Priority.toString());
                                  int j1939ExtendedDataPage = multiplexedIPDU.getJ1939ExtendedDataPage() ? 1 : 0;
                                  int j1939DataPage = multiplexedIPDU.getJ1939DataPage() ? 1 : 0;
                                  int j1939PduFormat = multiplexedIPDU.getJ1939PduFormat();
                                  int j1939PduSpecific = multiplexedIPDU.getJ1939PduSpecific();
                                  // 转二进制
                                  String j1939PriorityBinary = Integer.toBinaryString(j1939PriorityInt);
                                  String j1939PriorityBinaryFormat = String.format("%03d", Integer.valueOf(j1939PriorityBinary));
                                  String j1939ExtendedDataPageBinary = Integer.toBinaryString(j1939ExtendedDataPage);
                                  String j1939DataPageBinary = Integer.toBinaryString(j1939DataPage);
                                  String j1939PduFormatBinary = Integer.toBinaryString(j1939PduFormat);
                                  String j1939PduFormatBinaryFormat = String.format("%08d", Integer.valueOf(j1939PduFormatBinary));
                                  String j1939PduSpecificBinary = Integer.toBinaryString(j1939PduSpecific);
                                  String j1939PduSpecificBinaryFormat = String.format("%08d", Integer.valueOf(j1939PduSpecificBinary));
                                  String nmNodeIdBinary = Integer.toBinaryString(nmNodeId);
                                  String nmNodeIdBinaryFormat = String.format("%08d", Integer.valueOf(nmNodeIdBinary));
                                  String messageId = j1939PriorityBinaryFormat + j1939ExtendedDataPageBinary + j1939DataPageBinary + j1939PduFormatBinaryFormat + j1939PduSpecificBinaryFormat
                                          + nmNodeIdBinaryFormat;
                                  if (messageId.length() == 29) {
                                    BigInteger bigInteger = new BigInteger(messageId, 2);
                                    String messageIdHex = bigInteger.toString(16);
                                    bodyJson.addProperty("messageId", messageIdHex.toUpperCase());
                                  }
                                }
                              }
                            } else {
                              boolean j1939PduFormatDefaultMarker = mSignalIPDU.getJ1939PduFormatDefaultMarker();
                              boolean j1939PduSpecificDefaultMarker = mSignalIPDU.getJ1939PduSpecificDefaultMarker();
                              MJ1939PriorityEnum j1939Priority = mSignalIPDU.getJ1939Priority();
                              if (j1939Priority != null && !j1939PduFormatDefaultMarker && !j1939PduSpecificDefaultMarker && !nameDefaultMarker) {
                                int j1939PriorityInt = Integer.valueOf(j1939Priority.toString());
                                int j1939ExtendedDataPage = mSignalIPDU.getJ1939ExtendedDataPage() ? 1 : 0;
                                int j1939DataPage = mSignalIPDU.getJ1939DataPage() ? 1 : 0;
                                int j1939PduFormat = mSignalIPDU.getJ1939PduFormat();
                                int j1939PduSpecific = mSignalIPDU.getJ1939PduSpecific();
                                // 转二进制
                                String j1939PriorityBinary = Integer.toBinaryString(j1939PriorityInt);
                                String j1939PriorityBinaryFormat = String.format("%03d", Integer.valueOf(j1939PriorityBinary));
                                String j1939ExtendedDataPageBinary = Integer.toBinaryString(j1939ExtendedDataPage);
                                String j1939DataPageBinary = Integer.toBinaryString(j1939DataPage);
                                String j1939PduFormatBinary = Integer.toBinaryString(j1939PduFormat);
                                String j1939PduFormatBinaryFormat = String.format("%08d", Integer.valueOf(j1939PduFormatBinary));
                                String j1939PduSpecificBinary = Integer.toBinaryString(j1939PduSpecific);
                                String j1939PduSpecificBinaryFormat = String.format("%08d", Integer.valueOf(j1939PduSpecificBinary));
                                String nmNodeIdBinary = Integer.toBinaryString(nmNodeId);
                                String nmNodeIdBinaryFormat = String.format("%08d", Integer.valueOf(nmNodeIdBinary));
                                String messageId = j1939PriorityBinaryFormat + j1939ExtendedDataPageBinary + j1939DataPageBinary + j1939PduFormatBinaryFormat + j1939PduSpecificBinaryFormat
                                        + nmNodeIdBinaryFormat;
                                if (messageId.length() == 29) {
                                  BigInteger bigInteger = new BigInteger(messageId, 2);
                                  String messageIdHex = bigInteger.toString(16);
                                  bodyJson.addProperty("messageId", messageIdHex.toUpperCase());
                                }
                              }
                            }
                          }
                        }
                      }
                    }

                    // Type Definition（Signal Type列）
                    ceshiBySignal(mSignal, bodyJson);

                    // Timing
                    List<MCyclicTiming> cyclicTimings = mSignalTransmission.getCyclicTimings();
                    for (MCyclicTiming mCyclicTiming : cyclicTimings) {
                      bodyJson.addProperty("signalTransmissionCycle", mCyclicTiming.getTimePeriod());
                    }
                    Collection<MEventControlledTiming> eventControlledTimings = mSignalTransmission.getEventControlledTimings();
                    for (MEventControlledTiming mEventControlledTiming : eventControlledTimings) {
                      int numberOfRepetitions = mEventControlledTiming.getNumberOfRepetitions();
                      if (numberOfRepetitions != 0) {
                        double repetitionPeriod = mEventControlledTiming.getRepetitionPeriod();
                        String signalTransmissionCycle = numberOfRepetitions + "*" + repetitionPeriod;
                        bodyJson.addProperty("signalTransmissionCycle", signalTransmissionCycle);
                      }
                    }

                    // Reserved、IsSelf、New列值
                    MLayoutPackageArtefactOwner layoutPackageArtefactOwner = mSignal.getLayoutPackageArtefactOwner();
                    if (layoutPackageArtefactOwner instanceof MLayoutPackage) {
                      MLayoutPackage layoutPackage = (MLayoutPackage) layoutPackageArtefactOwner;
                      if (layoutPackage.getName().contains("Reserved")) {
                        bodyJson.addProperty("reserved", "r");
                      } else if (layoutPackage.getName().contains("IsSelf")) {
                        bodyJson.addProperty("isSelf", "r");
                      } else if (layoutPackage.getName().contains("New")) {
                        bodyJson.addProperty("new", "r");
                      }
                    }

                    jsonArray.add(bodyJson);
                  }
                }
              }
            }
          }
        }
      }
    }
    //其中构造方法的参数:
    //sortItem是要排序的jsonArray中一个元素, 这里我选择是Name, 也可以选择No或者是Length
    //sortType是排序的类型, 有三种情况
    // 1. 排序的元素对应的值是int， 那么sortType = "int";
    // 2. 排序的元素对应的值是string， 那么sortType = "string";
    // 3. 排序的元素对应的是是其他类型, 默认是不排序, (后面可以扩展)
    //sortDire是排序的方向, 可以是asc或者desc, 默认是数据的原始方向(就是没有排序方向)
    // 按照start bit排序
    JsonArray sortJsonArrayByStartBit = jsonArraySort(jsonArray, "startBit", "int", "asc");

    // 设置start bit
    if (sortJsonArrayByStartBit.size() > 0) {
      for (int j = 0; j < sortJsonArrayByStartBit.size(); j++) {
        JsonObject asJsonObject = sortJsonArrayByStartBit.get(j).getAsJsonObject();
        if (j == 0) {
          // 第一行
          asJsonObject.addProperty("startBit", "start bit");
        } else {
          // 如果是0，设置为""
          if (0 == asJsonObject.get("startBit").getAsInt()) {
            asJsonObject.addProperty("startBit", "");
          }
        }
      }
    }

    // 按照abbreviation排序
    JsonArray sortJsonArrayByAbbreviation = jsonArraySort(sortJsonArrayByStartBit, "abbreviation", "String", "asc");

    // abbreviation列空的放到最后
    JsonArray allArray = buildAbbreviationSortArray(sortJsonArrayByAbbreviation);

    if (transmissionStatus) {
      jsonObject.add("sheetInfo", allArray);
    } else {
      jsonObject.add("sheetInfo", new JsonArray());
    }
    jsonAllArray.add(jsonObject);
  }

  private String buildFrameData(JsonObject bodyJson, List<MPDUFrameAssignment> frameAssignments, MBusSystem mBusSystem, boolean isCanType) {
    if (frameAssignments != null && !frameAssignments.isEmpty()) {
      for (MPDUFrameAssignment mpduFrameAssignment : frameAssignments) {
        MFrame frame = mpduFrameAssignment.getFrame();
        if (frame instanceof MCANFrame) {
          MCANFrame mcanFrame = (MCANFrame) frame;
          List<MFrameTransmission> frameTransmissions = mcanFrame.getFrameTransmissions();
          for (MFrameTransmission mFrameTransmission : frameTransmissions) {
            if (mFrameTransmission instanceof MCANFrameTransmission) {
              MCANFrameTransmission canFrameTransmission = (MCANFrameTransmission) mFrameTransmission;
              MAbstractChannelCommunicationContentOwner frameTransmissionContainer = canFrameTransmission.getFrameTransmissionContainer();
              if (frameTransmissionContainer instanceof MChannelCommunication) {
                MChannelCommunication channelCommunication = (MChannelCommunication) frameTransmissionContainer;
                MAbstractBusRoutingArtefactOwner busRoutingArtefactOwner = channelCommunication.getBusRoutingArtefactOwner();
                if (busRoutingArtefactOwner instanceof MCANCommunication) {
                  MCANCommunication canCommunication = (MCANCommunication) busRoutingArtefactOwner;
                  MSignalConnection signalConnection = canCommunication.getSignalConnection();
                  if (signalConnection instanceof MBusSystem) {
                    MBusSystem busSystem = (MBusSystem) signalConnection;
                    if (busSystem == mBusSystem) {
                      // General
                      String abbreviation = mcanFrame.getName();
                      String messageName = mcanFrame.getLongName();
                      int messageLength = mcanFrame.getByteLength();

                      String canFrameName = abbreviation;
                      // 查询ecu
                      String ecuName = loopGetEcuName(canFrameTransmission);
                      if (StringUtils.isNotBlank(ecuName) && isCanType) {
                        canFrameName = ecuName + "_" + abbreviation;
                      }
                      bodyJson.addProperty("abbreviation", canFrameName);
                      bodyJson.addProperty("messageName", StringUtils.isBlank(messageName) ? "" : messageName);
                      bodyJson.addProperty("messageLength", messageLength);
                      // Type Definition Set
                      ceshi(canFrameTransmission, bodyJson);

                      // 获取Nm Node ID（sending）
                      int nmNodeId = 0;
                      boolean nameDefaultMarker = true;
                      Collection<MNmNodeCommunication> nmNodeCommunications = canFrameTransmission.getSendingNmNode();
                      for (MNmNodeCommunication nmNodeCommunication : nmNodeCommunications) {
                        nmNodeId = nmNodeCommunication.getNmNodeId();
                        nameDefaultMarker = nmNodeCommunication.getNmNodeIdDefaultMarker();
                      }
                      return abbreviation + "!" + nmNodeId + "!" + nameDefaultMarker;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
    return "";
  }

  private String loopGetEcuName(MCANFrameTransmission canFrameTransmission) {
    String ecuName = "";
    Collection<MFrameGatewayRoutingEntry> routingEntriesOut = canFrameTransmission.getRoutingEntriesOut();
    if (routingEntriesOut.isEmpty()) {
      Collection<MECUInterface> sendingECUInterface = canFrameTransmission.getSendingECUInterface();
      for (MECUInterface mecuInterface : sendingECUInterface) {
        if (mecuInterface instanceof MCANECUInterface) {
          MCANECUInterface canecuInterface = (MCANECUInterface) mecuInterface;
          MSignalConnector connector = canecuInterface.getConnector();
          if (connector instanceof MBusConnector) {
            MBusConnector busConnector = (MBusConnector) connector;
            MECU mecu = (MECU) busConnector.getElectronicComposite();
            ecuName = mecu.getName();
          }
        }
      }
    } else {
      for (MFrameGatewayRoutingEntry mFrameGatewayRoutingEntry : routingEntriesOut) {
        MFrameTransmission inTransmissions = mFrameGatewayRoutingEntry.getInTransmission();
        if (inTransmissions instanceof MCANFrameTransmission) {
          MCANFrameTransmission canFrameTransmission1 = (MCANFrameTransmission) inTransmissions;
          loopGetEcuName(canFrameTransmission1);
        }
      }
    }
    return ecuName;
  }

  private JsonArray buildAbbreviationSortArray(JsonArray sortJsonArrayByAbbreviation) {
    JsonArray nullArray = new JsonArray();
    JsonArray notNullArray = new JsonArray();
    for (int i = 0; i < sortJsonArrayByAbbreviation.size(); i++) {
      JsonObject jsonObj = sortJsonArrayByAbbreviation.get(i).getAsJsonObject();
      if (StringUtils.isBlank(jsonObj.get("abbreviation").getAsString())) {
        nullArray.add(jsonObj);
      } else {
        notNullArray.add(jsonObj);
      }
    }
    notNullArray.addAll(nullArray);
    return notNullArray;
  }

  private JsonObject initFrontHeaderJsonData() {
    JsonObject headerJson = new JsonObject();
    headerJson.addProperty("messageId", "Message ID");
    headerJson.addProperty("abbreviation", "Abbreviation");
    headerJson.addProperty("messageName", "Message Name");
    headerJson.addProperty("cyclic", "cyclic[ms]");
    headerJson.addProperty("sendType", "Send Type");
    headerJson.addProperty("messageLength", "Message Length [Byte]");
    headerJson.addProperty("multiplexingValue", "Multiplexing/Value");
    headerJson.addProperty("byteNumber", "Byte Number");
    headerJson.addProperty("bitNumber", "Bit Number");
    headerJson.addProperty("signalLength", "Signal Length [Bit]");
    headerJson.addProperty("startBit", "start bit");
    headerJson.addProperty("signalName", "Signal Name");
    headerJson.addProperty("signalDescription", "Signal Description");
    headerJson.addProperty("signalDefault", "Signal Default");
    return headerJson;
  }

  private JsonObject initAfterHeaderJsonData(JsonObject headerJson, String excelName) {
    String name = excelName.replace("_", "").replace("CAN", "");
    headerJson.addProperty("reserved", "Reserved_" + name);
    headerJson.addProperty("isSelf", "IsSelf_" + name);
    headerJson.addProperty("new", "New_" + name);
    headerJson.addProperty("physicalRange", "Physical Range");
    headerJson.addProperty("normal", "Normal");
    headerJson.addProperty("resolution", "Resolution");
    headerJson.addProperty("cycleTimeFast", "Cycle Time Fast(ms)");
    headerJson.addProperty("nrOfReption", "Nr.Of Reption");
    headerJson.addProperty("delayTime", "Delay Time(ms)");
    headerJson.addProperty("spn", "SPN");
    headerJson.addProperty("signalType", "Signal Type");
    headerJson.addProperty("signalTransmissionCycle", "Signal transmission cycle");
    return headerJson;
  }

  private JsonObject initBodyJsonData(List<String> ecuList) {
    JsonObject bodyJson = new JsonObject();
    bodyJson.addProperty("messageId", "");
    bodyJson.addProperty("abbreviation", "");
    bodyJson.addProperty("messageName", "");
    bodyJson.addProperty("cyclic", "");
    bodyJson.addProperty("sendType", "");
    bodyJson.addProperty("messageLength", "");
    bodyJson.addProperty("multiplexingValue", "");
    bodyJson.addProperty("byteNumber", "");
    bodyJson.addProperty("bitNumber", "");
    bodyJson.addProperty("signalLength", 0);
    bodyJson.addProperty("startBit", "");
    bodyJson.addProperty("signalName", "");
    bodyJson.addProperty("signalDescription", "");
    bodyJson.addProperty("signalDefault", "");
    for (String ecuName : ecuList) {
      bodyJson.addProperty(ecuName, "");
    }
    bodyJson.addProperty("reserved", "");
    bodyJson.addProperty("isSelf", "");
    bodyJson.addProperty("new", "");
    bodyJson.addProperty("physicalRange", "");
    bodyJson.addProperty("normal", "");
    bodyJson.addProperty("resolution", "");
    bodyJson.addProperty("cycleTimeFast", "");
    bodyJson.addProperty("nrOfReption", "");
    bodyJson.addProperty("delayTime", "");
    bodyJson.addProperty("spn", 0);
    bodyJson.addProperty("signalType", "");
    bodyJson.addProperty("signalTransmissionCycle", "");
    bodyJson.addProperty("type", "");
    bodyJson.addProperty("table", "");
    bodyJson.addProperty("isisMultiplexor", "");
    return bodyJson;
  }

  private void buildLinerConversion(MLinearConversion mLinearConversion, JsonObject jObject) {
    //    warning("---------------Liner Conversion---------------");
    String min = mLinearConversion.getMin();
    String max = mLinearConversion.getMax();
    double physicalResolution = mLinearConversion.getFactor();
    if (StringUtils.endsWith(min, ".0")) {
      min = min.substring(0, min.length() - 2);
    }
    if (StringUtils.endsWith(max, ".0")) {
      max = max.substring(0, max.length() - 2);
    }
    jObject.addProperty("normal", min + ".." + max);
    jObject.addProperty("resolution", physicalResolution);

    jObject.addProperty("type", "liner");
    jObject.addProperty("table", "");
  }

  private void buildLinearVerbalTableConversion(MLinearVerbalTableConversion linearVerbalTableConversion, JsonObject jObject) {
    //    warning("---------------Liner Conversion---------------");
    String min = linearVerbalTableConversion.getMin();
    String max = linearVerbalTableConversion.getMax();
    double physicalResolution = linearVerbalTableConversion.getFactor();
    if (StringUtils.endsWith(min, ".0")) {
      min = min.substring(0, min.length() - 2);
    }
    if (StringUtils.endsWith(max, ".0")) {
      max = max.substring(0, max.length() - 2);
    }
    jObject.addProperty("normal", min + ".." + max);
    jObject.addProperty("resolution", physicalResolution);

    jObject.addProperty("type", "liner");
    jObject.addProperty("table", "");
  }

  private void buildVerbalTableConversion(List<MTableValue> tableValues, JsonObject jObject, String signalName) {
    //    warning("---------------Verbal Table Conversion---------------");
    int signalLength = jObject.get("signalLength").getAsInt();
    String multiplexingValue = jObject.get("multiplexingValue").getAsString();
    if ("Multiplexor".equals(multiplexingValue)) {
      signalLength = tableValues.size() / 2;
    } else {
      if (signalLength > 8) {
        logErrorMessage("错误：离散型signalLength长度超过8，错误signal= " + jObject.get("signalName").getAsString());
        return;
      }
    }

    // 第一行
    JsonArray jsonArray = new JsonArray();
    JsonObject jsonObject = new JsonObject();
    for (int i = signalLength - 1; i >= 0; i--) {
      jsonObject.addProperty("b" + i, "b" + i);
    }
    jsonObject.addProperty("size", signalLength);
    jsonObject.addProperty("valueDescription", "Function");
    jsonArray.add(jsonObject);

    // 第二行开始
    // 新集合，用于保存根据xIndex重新排序
    JsonArray jsonArraySort = new JsonArray();
    if ("Multiplexor".equals(multiplexingValue)) {
      for (MTableValue mTableValue : tableValues) {
        String valueDescription = mTableValue.getName();
        String xIndex = mTableValue.getXIndex();
        try {
          Integer xIndexInt = Integer.valueOf(xIndex);
          JsonObject jsonObjectSort = new JsonObject();
          jsonObjectSort.addProperty("xIndex", xIndexInt);
          jsonObjectSort.addProperty("valueDescription", valueDescription);
          jsonArraySort.add(jsonObjectSort);
        } catch (Exception e) {
          logErrorMessage("错误：信号（" + signalName + "）的xIndex值不是数字！");
        }
      }
    } else {
      for (int i = 0; i < Math.pow(2, signalLength); i++) {
        boolean hasXIndex = false;
        JsonObject jsonObjectSort = new JsonObject();
        for (MTableValue mTableValue : tableValues) {
          String xIndex = mTableValue.getXIndex();
          try {
            Integer xIndexInt = Integer.valueOf(xIndex);
            if (i == xIndexInt) {
              String valueDescription = mTableValue.getName();
              jsonObjectSort.addProperty("xIndex", xIndexInt);
              jsonObjectSort.addProperty("valueDescription", valueDescription);
              jsonArraySort.add(jsonObjectSort);
              hasXIndex = true;
              break;
            }
          } catch (Exception e) {
            logErrorMessage("错误：信号（" + signalName + "）的xIndex值不是数字！");
          }
        }
        if (!hasXIndex) {
          jsonObjectSort.addProperty("xIndex", i);
          jsonObjectSort.addProperty("valueDescription", "Reserved");
          jsonArraySort.add(jsonObjectSort);
        }
      }
    }
    // 按照xIndex升序排序
    JsonArray newJsonArray = jsonArraySort(jsonArraySort, "xIndex", "int", "asc");

    if (tableValues.size() > 2) {
      int nextIndex = 0;
      for (int i = 0; i < newJsonArray.size(); i++) {
        JsonObject asJsonObject = newJsonArray.get(i).getAsJsonObject();
        int xIndex = asJsonObject.get("xIndex").getAsInt();
        String valueDescription = asJsonObject.get("valueDescription").getAsString();

        if (0 < i && (i + 2) < newJsonArray.size()) { // 1 --> 倒是第3个
          String valueDescriptionOld = newJsonArray.get(i - 1).getAsJsonObject().get("valueDescription").getAsString();
          String valueDescriptionNext = newJsonArray.get(i + 1).getAsJsonObject().get("valueDescription").getAsString();
          String valueDescriptionNextNext = newJsonArray.get(i + 2).getAsJsonObject().get("valueDescription").getAsString();

          if (valueDescription.equals(valueDescriptionOld) && valueDescription.equals(valueDescriptionNext) && !valueDescription.equals(valueDescriptionNextNext)) {
            buildEllipsisTable(signalLength, jsonArray);
          }
          if (!valueDescription.equals(valueDescriptionOld) || !valueDescription.equals(valueDescriptionNext)) {
            buildTable(signalLength, jsonArray, xIndex, valueDescription);
          }
        } else if ((i + 1) == newJsonArray.size() - 1) { // 倒数第2个
          String valueDescriptionOld = newJsonArray.get(i - 1).getAsJsonObject().get("valueDescription").getAsString();
          String valueDescriptionNext = newJsonArray.get(i + 1).getAsJsonObject().get("valueDescription").getAsString();
          if (valueDescription.equals(valueDescriptionOld) && valueDescription.equals(valueDescriptionNext)) {
            buildEllipsisTable(signalLength, jsonArray);
          }
          if (!valueDescription.equals(valueDescriptionOld) || !valueDescription.equals(valueDescriptionNext)) {
            buildTable(signalLength, jsonArray, xIndex, valueDescription);
          }
        } else { // 0 和 最后1个
          if (nextIndex != xIndex) {
            buildEllipsisTable(signalLength, jsonArray);
          }
          buildTable(signalLength, jsonArray, xIndex, valueDescription);
        }

        nextIndex = xIndex + 1;
      }
    } else {
      for (int i = 0; i < newJsonArray.size(); i++) {
        JsonObject asJsonObject = newJsonArray.get(i).getAsJsonObject();
        int xIndex = asJsonObject.get("xIndex").getAsInt();
        String valueDescription = asJsonObject.get("valueDescription").getAsString();
        buildTable(signalLength, jsonArray, xIndex, valueDescription);
      }
    }

    jObject.addProperty("type", "table");
    jObject.add("table", jsonArray);
  }

  private void buildTable(int signalLength, JsonArray jsonArray, int xIndex, String valueDescription) {
    String xIndexBinary = Integer.toBinaryString(xIndex);
    String sizeString = String.format("%0" + signalLength + "d", Integer.valueOf(xIndexBinary));
    StringBuilder sb = new StringBuilder(sizeString);
    String[] arr = sb.reverse().toString().split("");

    JsonObject bodyJson = new JsonObject();
    for (int j = signalLength - 1; j >= 0; j--) {
      bodyJson.addProperty("b" + j, arr[j]);
    }
    bodyJson.addProperty("valueDescription", valueDescription);
    jsonArray.add(bodyJson);
  }

  private void buildEllipsisTable(int signalLength, JsonArray jsonArray) {
    JsonObject bodyJson = new JsonObject();
    for (int j = signalLength - 1; j >= 0; j--) {
      bodyJson.addProperty("b" + j, "");
    }
    bodyJson.addProperty("valueDescription", "...");
    jsonArray.add(bodyJson);
  }

  // 自定义属性
  private void ceshi(MCANFrameTransmission canFrameTransmission, JsonObject jObject) {
    //获取当前构件类型所有自定义属性类型
    List<MMetricAttributeDefinition> allValidAttributeDefinitions = AttributeDefinitionUtility.getAllValidAttributeDefinitions(canFrameTransmission);
    //    warning("allValidAttributeDefinitions：" + allValidAttributeDefinitions);
    for (MMetricAttributeDefinition mMetricAttributeDefinition : allValidAttributeDefinitions) {
      //获取该指定构件 某个自定义属性的值
      if (mMetricAttributeDefinition.getName().equals("GenMsgCycleTime")) {
        MAbstractAttributeValue abstractAttributeValue = AttributeDefinitionUtility.getAttributeValue(canFrameTransmission, mMetricAttributeDefinition);
        String value = getValByType(abstractAttributeValue);
        jObject.addProperty("cyclic", value);
      } else if (mMetricAttributeDefinition.getName().equals("GenMsgSendType")) {
        MAbstractAttributeValue abstractAttributeValue = AttributeDefinitionUtility.getAttributeValue(canFrameTransmission, mMetricAttributeDefinition);
        String value = getValByType(abstractAttributeValue);
        if (StringUtils.isNotBlank(value)) {
          String[] sendTypeArr = value.split(":");
          if (sendTypeArr.length == 2) {
            jObject.addProperty("sendType", sendTypeArr[1].trim());
          }
        }
      } else if (mMetricAttributeDefinition.getName().equals("GenMsgCycleTimeFast")) {
        MAbstractAttributeValue abstractAttributeValue = AttributeDefinitionUtility.getAttributeValue(canFrameTransmission, mMetricAttributeDefinition);
        String value = getValByType(abstractAttributeValue);
        jObject.addProperty("cycleTimeFast", value);
      } else if (mMetricAttributeDefinition.getName().equals("GenMsgNrOfRepetition")) {
        MAbstractAttributeValue abstractAttributeValue = AttributeDefinitionUtility.getAttributeValue(canFrameTransmission, mMetricAttributeDefinition);
        String value = getValByType(abstractAttributeValue);
        jObject.addProperty("nrOfReption", value);
      } else if (mMetricAttributeDefinition.getName().equals("GenMsgDelayTime")) {
        MAbstractAttributeValue abstractAttributeValue = AttributeDefinitionUtility.getAttributeValue(canFrameTransmission, mMetricAttributeDefinition);
        String value = getValByType(abstractAttributeValue);
        jObject.addProperty("delayTime", value);
      }
    }
  }

  //自定义属性
  private void ceshiBySignal(MSignal signal, JsonObject jObject) {
    //获取当前构件类型所有自定义属性类型
    List<MMetricAttributeDefinition> allValidAttributeDefinitions = AttributeDefinitionUtility.getAllValidAttributeDefinitions(signal);
    //    warning("allValidAttributeDefinitions：" + allValidAttributeDefinitions);
    for (MMetricAttributeDefinition mMetricAttributeDefinition : allValidAttributeDefinitions) {
      //获取该指定构件 某个自定义属性的值
      if (mMetricAttributeDefinition.getName().equals("GenSigSendType")) {
        MAbstractAttributeValue abstractAttributeValue = AttributeDefinitionUtility.getAttributeValue(signal, mMetricAttributeDefinition);
        String value = getValByType(abstractAttributeValue);
        if (StringUtils.isNotBlank(value)) {
          if (value.contains(": ")) {
            jObject.addProperty("signalType", value.split(": ")[1]);
          } else {
            logErrorMessage("错误：信号（" + signal.getName() + "）设置的自定义属性GenSigSendType值格式不正确！");
          }
        } else {
          logErrorMessage("错误：信号（" + signal.getName() + "）没有设置自定义属性GenSigSendType值！");
        }
      }
    }
  }

  private String getValByType(MAbstractAttributeValue abstractAttributeValue) {
    if (abstractAttributeValue != null) {
      if (abstractAttributeValue instanceof MLightWeightStringAttributeValue) {
        MLightWeightStringAttributeValue attributeValue = (MLightWeightStringAttributeValue) abstractAttributeValue;
        return attributeValue.getLwValue();
      } else if (abstractAttributeValue instanceof MLightWeightEnumAttributeValue) {
        MLightWeightEnumAttributeValue attributeValue = (MLightWeightEnumAttributeValue) abstractAttributeValue;
        MAbstractEnumEntry usedEnumEntry = attributeValue.getUsedEnumEntry();
        if (usedEnumEntry instanceof MStringLiteral) {
          return ((MStringLiteral) usedEnumEntry).getName();
        }
      } else if (abstractAttributeValue instanceof MLightWeightIntegerAttributeValue) {
        MLightWeightIntegerAttributeValue attributeValue = (MLightWeightIntegerAttributeValue) abstractAttributeValue;
        return String.valueOf(attributeValue.getLwValue());
      } else if (abstractAttributeValue instanceof MLightWeightByteValue) {
        MLightWeightByteValue attributeValue = (MLightWeightByteValue) abstractAttributeValue;
        return String.valueOf(attributeValue.getLwValue());
      } else if (abstractAttributeValue instanceof MLightWeightShortAttributeValue) {
        MLightWeightShortAttributeValue attributeValue = (MLightWeightShortAttributeValue) abstractAttributeValue;
        return String.valueOf(attributeValue.getLwValue());
      } else if (abstractAttributeValue instanceof MLightWeightBooleanAttributeValue) {
        MLightWeightBooleanAttributeValue attributeValue = (MLightWeightBooleanAttributeValue) abstractAttributeValue;
        return String.valueOf(attributeValue.getLwValue());
      }
    }
    //  MLightWeightByteValue
    return "";
  }

  /**
   * json排序
   */
  private JsonArray jsonArraySort(JsonArray jsonArray, String sortItem, String sortType, String sortDire) {
    //    warning("sortItem= " + sortItem);
    //    warning("init jsonArray= " + jsonArray.toString());

    //这里最核心的地方就是SortComparator这个类
    JsonArray sort_JsonArray = new JsonArray();
    List<JsonObject> list = new ArrayList<>();
    JsonObject headlerJsonObj = new JsonObject();
    for (int i = 0; i < jsonArray.size(); i++) {
      JsonObject jsonObj = (JsonObject) jsonArray.get(i);
      if (jsonObj.has("abbreviation") && "Abbreviation".equals(jsonObj.get("abbreviation").getAsString())) {
        // 获取标题行
        headlerJsonObj = jsonObj;
      } else {
        if (jsonObj.get(sortItem).getAsJsonPrimitive().isString() && "int".equals(sortType)) {
          jsonObj.addProperty(sortItem, 0);
        }
        list.add(jsonObj);
      }
    }
    list.sort(new SortComparator202302141(sortItem, sortType, sortDire));

    // 将标题行放到新JsonArray中第一行
    if (headlerJsonObj.size() != 0) {
      sort_JsonArray.add(headlerJsonObj);
    }
    for (JsonObject jsonObject : list) {
      sort_JsonArray.add(jsonObject);
    }
    //    warning("after sort_JsonArray= " + sort_JsonArray);
    return sort_JsonArray;
  }
}

class SortComparator202302141 implements Comparator<JsonObject> {

  private final String sortItem;

  private final String sortType;

  private final String sortDire;

  @SuppressWarnings("hiding")
  public SortComparator202302141(String sortItem, String sortType, String sortDire) {
    this.sortItem = sortItem;
    this.sortType = sortType;
    this.sortDire = sortDire;
  }

  @SuppressWarnings("nls")
  @Override
  public int compare(JsonObject o1, JsonObject o2) {
    String value1 = o1.getAsJsonObject().get(sortItem).getAsString();
    String value2 = o2.getAsJsonObject().get(sortItem).getAsString();
    if ("int".equalsIgnoreCase(sortType)) { // int sort
      int int1 = Integer.parseInt(value1);
      int int2 = Integer.parseInt(value2);
      if ("asc".equalsIgnoreCase(sortDire)) {
        return int1 - int2;
      } else if ("desc".equalsIgnoreCase(sortDire)) {
        return int2 - int1;
      } else {
        return 0;
      }
    } else if ("string".equalsIgnoreCase(sortType)) { // string sort
      if ("asc".equalsIgnoreCase(sortDire)) {
        return value1.compareTo(value2);
      } else if ("desc".equalsIgnoreCase(sortDire)) {
        return value2.compareTo(value1);
      } else {
        return 0;
      }
    } else { // nothing sort
      return 0;
    }
  }

}
