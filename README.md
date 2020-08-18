# poi-export-word
java使用poi操作word, 支持动态的行(一个占位符插入多条)和表格中动态行, 支持图片)
博客地址：https://blog.csdn.net/qq_37880968/article/details/102870963

## 模板图
希望可以帮到大家，希望给个start
[项目git源码地址](https://github.com/lbxgit/poi-export-word.git)
![在这里插入图片描述](https://img-blog.csdnimg.cn/20191127113817917.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzM3ODgwOTY4,size_16,color_FFFFFF,t_70)
## 效果图
![在这里插入图片描述](https://img-blog.csdnimg.cn/20191127113727737.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzM3ODgwOTY4,size_16,color_FFFFFF,t_70)

## 1，引入maven依赖

```
<dependency>
        <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>3.17</version>
</dependency>

<dependency>
        <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>3.17</version>
</dependency>

<dependency>
        <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml-schemas</artifactId>
    <version>3.17</version>
</dependency>
```
## 2，核心工具类
```
package per.qiao.utils.hutool.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.junit.Assert;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * @author: qyp
 * Date: 2019-10-25 14:48
 */
public class DynWordUtils {

    private final Logger logger = LoggerFactory.getLogger(DynWordUtils.class);

    /**
     * 被list替换的段落 被替换的都是oldParagraph
     */
    private XWPFParagraph oldParagraph;

    /**
     * 参数
     */
    private Map<String, Object> paramMap;

    /**
     * 当前元素的位置
     */
    int n = 0;

    /**
     * 判断当前是否是遍历的表格
     */
    boolean isTable = false;

    /**
     * 模板对象
     */
    XWPFDocument templateDoc;

    /**
     * 默认字体的大小
     */
    final int DEFAULT_FONT_SIZE = 10;

    /**
     * 重复模式的占位符所在的行索引
     */
    private int currentRowIndex;

    /**
     * 入口
     *
     * @param paramMap     模板中使用的参数
     * @param templatePaht 模板全路径
     * @param outPath      生成的文件存放的本地全路径
     */
    public static void process(Map<String, Object> paramMap, String templatePaht, String outPath) {
        DynWordUtils dynWordUtils = new DynWordUtils();
        dynWordUtils.setParamMap(paramMap);
        dynWordUtils.createWord(templatePaht, outPath);
    }

    /**
     * 生成动态的word
     * @param templatePath
     * @param outPath
     */
    public void createWord(String templatePath, String outPath) {
        File inFile = new File(templatePath);
        try (FileOutputStream outStream = new FileOutputStream(outPath)) {
            templateDoc = new XWPFDocument(OPCPackage.open(inFile));
            parseTemplateWord();
            templateDoc.write(outStream);
        } catch (Exception e) {
            StackTraceElement[] stackTrace = e.getStackTrace();

            String className = stackTrace[0].getClassName();
            String methodName = stackTrace[0].getMethodName();
            int lineNumber = stackTrace[0].getLineNumber();

            logger.error("错误：第:{}行, 类名:{}, 方法名:{}", lineNumber, className, methodName);
            throw new RuntimeException(e.getCause().getMessage());
        }
    }

    /**
     * 解析word模板
     */
    public void parseTemplateWord() throws Exception {

        List<IBodyElement> elements = templateDoc.getBodyElements();

        for (; n < elements.size(); n++) {
            IBodyElement element = elements.get(n);
            // 普通段落
            if (element instanceof XWPFParagraph) {

                XWPFParagraph paragraph = (XWPFParagraph) element;
                oldParagraph = paragraph;
                if (paragraph.getParagraphText().isEmpty()) {
                    continue;
                }

                delParagraph(paragraph);

            } else if (element instanceof XWPFTable) {
                // 表格
                isTable = true;
                XWPFTable table = (XWPFTable) element;

                delTable(table, paramMap);
                isTable = false;
            }
        }

    }

    /**
     * 处理段落
     */
    private void delParagraph(XWPFParagraph paragraph) throws Exception {
        List<XWPFRun> runs = oldParagraph.getRuns();
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text == null) {
                continue;
            }
            sb.append(text);
            run.setText("", 0);
        }
        Placeholder(paragraph, runs, sb);
    }


    /**
     * 匹配传入信息集合与模板
     *
     * @param placeholder 模板需要替换的区域()
     * @param paramMap    传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public void changeValue(XWPFRun currRun, String placeholder, Map<String, Object> paramMap) throws Exception {

        String placeholderValue = placeholder;
        if (paramMap == null || paramMap.isEmpty()) {
            return;
        }

        Set<Map.Entry<String, Object>> textSets = paramMap.entrySet();
        for (Map.Entry<String, Object> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String mapKey = textSet.getKey();
            String docKey = PoiWordUtils.getDocKey(mapKey);

            if (placeholderValue.indexOf(docKey) != -1) {
                Object obj = textSet.getValue();
                // 需要添加一个list
                if (obj instanceof List) {
                    placeholderValue = delDynList(placeholder, (List) obj);
                } else {
                    placeholderValue = placeholderValue.replaceAll(
                            PoiWordUtils.getPlaceholderReg(mapKey)
                            , String.valueOf(obj));
                }
            }
        }

        currRun.setText(placeholderValue, 0);
    }

    /**
     * 处理的动态的段落（参数为list）
     *
     * @param placeholder 段落占位符
     * @param obj
     * @return
     */
    private String delDynList(String placeholder, List obj) {
        String placeholderValue = placeholder;
        List dataList = obj;
        Collections.reverse(dataList);
        for (int i = 0, size = dataList.size(); i < size; i++) {
            Object text = dataList.get(i);
            // 占位符的那行, 不用重新创建新的行
            if (i == 0) {
                placeholderValue = String.valueOf(text);
            } else {
                XWPFParagraph paragraph = createParagraph(String.valueOf(text));
                if (paragraph != null) {
                    oldParagraph = paragraph;
                }
                // 增加段落后doc文档会的element的size会随着增加（在当前行的上面添加），回退并解析新增的行（因为可能新增的带有占位符，这里为了支持图片和表格）
                if (!isTable) {
                    n--;
                }
            }
        }
        return placeholderValue;
    }

    /**
     * 创建段落 <p></p>
     *
     * @param texts
     */
    public XWPFParagraph createParagraph(String... texts) {

        // 使用游标创建一个新行
        XmlCursor cursor = oldParagraph.getCTP().newCursor();
        XWPFParagraph newPar = templateDoc.insertNewParagraph(cursor);
        // 设置段落样式
        newPar.getCTP().setPPr(oldParagraph.getCTP().getPPr());

        copyParagraph(oldParagraph, newPar, texts);

        return newPar;
    }

    /**
     * 处理表格（遍历）
     *
     * @param table    表格
     * @param paramMap 需要替换的信息集合
     */
    public void delTable(XWPFTable table, Map<String, Object> paramMap) throws Exception {
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0, size = rows.size(); i < size; i++) {
            XWPFTableRow row = rows.get(i);
            currentRowIndex = i;
            // 如果是动态添加行 直接处理后返回
            if (delAndJudgeRow(table, paramMap, row)) {
                return;
            }
        }
    }

    /**
     * 判断并且是否是动态行，并且处理表格占位符
     * @param table 表格对象
     * @param paramMap 参数map
     * @param row 当前行
     * @return
     * @throws Exception
     */
    private boolean delAndJudgeRow(XWPFTable table, Map<String, Object> paramMap, XWPFTableRow row) throws Exception {
        // 当前行是动态行标志
        if (PoiWordUtils.isAddRow(row)) {
            List<XWPFTableRow> xwpfTableRows = addAndGetRows(table, row, paramMap);
            // 回溯添加的行，这里是试图处理动态添加的图片
            for (XWPFTableRow tbRow : xwpfTableRows) {
                delAndJudgeRow(table, paramMap, tbRow);
            }
            return true;
        }

        // 如果是重复添加的行
        if (PoiWordUtils.isAddRowRepeat(row)) {
            List<XWPFTableRow> xwpfTableRows = addAndGetRepeatRows(table, row, paramMap);
            // 回溯添加的行，这里是试图处理动态添加的图片
            for (XWPFTableRow tbRow : xwpfTableRows) {
                delAndJudgeRow(table, paramMap, tbRow);
            }
            return true;
        }
        // 当前行非动态行标签
        List<XWPFTableCell> cells = row.getTableCells();
        for (XWPFTableCell cell : cells) {
            //判断单元格是否需要替换
            if (PoiWordUtils.checkText(cell.getText())) {
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    List<XWPFRun> runs = paragraph.getRuns();
                    StringBuilder sb = new StringBuilder();
                    for (XWPFRun run : runs) {
                        sb.append(run.toString());
                        run.setText("", 0);
                    }
                    Placeholder(paragraph, runs, sb);
                }
            }
        }
        return false;
    }

    /**
     * 处理占位符
     * @param runs 当前段的runs
     * @param sb 当前段的内容
     * @throws Exception
     */
    private void Placeholder(XWPFParagraph currentPar, List<XWPFRun> runs, StringBuilder sb) throws Exception {
        if (runs.size() > 0) {
            String text = sb.toString();
            XWPFRun currRun = runs.get(0);
            if (PoiWordUtils.isPicture(text)) {
                // 该段落是图片占位符
                ImageEntity imageEntity = (ImageEntity) PoiWordUtils.getValueByPlaceholder(paramMap, text);
                int indentationFirstLine = currentPar.getIndentationFirstLine();
                // 清除段落的格式，否则图片的缩进有问题
                currentPar.getCTP().setPPr(null);
                //设置缩进
                currentPar.setIndentationFirstLine(indentationFirstLine);
                addPicture(currRun, imageEntity);
            } else {
                changeValue(currRun, text, paramMap);
            }
        }
    }

    /**
     * 添加图片
     * @param currRun 当前run
     * @param imageEntity 图片对象
     * @throws InvalidFormatException
     * @throws FileNotFoundException
     */
    private void addPicture(XWPFRun currRun, ImageEntity imageEntity) throws InvalidFormatException, FileNotFoundException {
        Integer typeId = imageEntity.getTypeId().getTypeId();
        String picId = currRun.getDocument().addPictureData(new FileInputStream(imageEntity.getUrl()), typeId);
        ImageUtils.createPicture(currRun, picId, templateDoc.getNextPicNameNumber(typeId),
                imageEntity.getWidth(), imageEntity.getHeight());
    }

    /**
     * 添加行  标签行不是新创建的
     *
     * @param table
     * @param flagRow  flagRow 表有标签的行
     * @param paramMap 参数
     */
    private List<XWPFTableRow> addAndGetRows(XWPFTable table, XWPFTableRow flagRow, Map<String, Object> paramMap) throws Exception {
        List<XWPFTableCell> flagRowCells = flagRow.getTableCells();
        XWPFTableCell flagCell = flagRowCells.get(0);

        String text = flagCell.getText();
        List<List<String>> dataList = (List<List<String>>) PoiWordUtils.getValueByPlaceholder(paramMap, text);

        // 新添加的行
        List<XWPFTableRow> newRows = new ArrayList<>(dataList.size());
        if (dataList == null || dataList.size() <= 0) {
            return newRows;
        }

        XWPFTableRow currentRow = flagRow;
        int cellSize = flagRow.getTableCells().size();
        for (int i = 0, size = dataList.size(); i < size; i++) {
            if (i != 0) {
                currentRow = table.createRow();
                // 复制样式
                if (flagRow.getCtRow() != null) {
                    currentRow.getCtRow().setTrPr(flagRow.getCtRow().getTrPr());
                }
            }
            addRow(flagCell, currentRow, cellSize, dataList.get(i));
            newRows.add(currentRow);
        }
        return newRows;
    }

    /**
     * 添加重复多行 动态行  每一行都是新创建的
     * @param table
     * @param flagRow
     * @param paramMap
     * @return
     * @throws Exception
     */
    private List<XWPFTableRow> addAndGetRepeatRows(XWPFTable table, XWPFTableRow flagRow, Map<String, Object> paramMap) throws Exception {
        List<XWPFTableCell> flagRowCells = flagRow.getTableCells();
        XWPFTableCell flagCell = flagRowCells.get(0);
        String text = flagCell.getText();
        List<List<String>> dataList = (List<List<String>>) PoiWordUtils.getValueByPlaceholder(paramMap, text);
        String tbRepeatMatrix = PoiWordUtils.getTbRepeatMatrix(text);
        Assert.assertNotNull("模板矩阵不能为空", tbRepeatMatrix);

        // 新添加的行
        List<XWPFTableRow> newRows = new ArrayList<>(dataList.size());
        if (dataList == null || dataList.size() <= 0) {
            return newRows;
        }

        String[] split = tbRepeatMatrix.split(PoiWordUtils.tbRepeatMatrixSeparator);
        int startRow = Integer.parseInt(split[0]);
        int endRow = Integer.parseInt(split[1]);
        int startCell = Integer.parseInt(split[2]);
        int endCell = Integer.parseInt(split[3]);

        XWPFTableRow currentRow;
        for (int i = 0, size = dataList.size(); i < size; i++) {
            int flagRowIndex = i % (endRow - startRow + 1);
            XWPFTableRow repeatFlagRow = table.getRow(flagRowIndex);
            // 清除占位符那行
            if (i == 0) {
                table.removeRow(currentRowIndex);
            }
            currentRow = table.createRow();
            // 复制样式
            if (repeatFlagRow.getCtRow() != null) {
                currentRow.getCtRow().setTrPr(repeatFlagRow.getCtRow().getTrPr());
            }
            addRowRepeat(startCell, endCell, currentRow, repeatFlagRow, dataList.get(i));
            newRows.add(currentRow);
        }
        return newRows;
    }

    /**
     * 根据模板cell添加新行
     *
     * @param flagCell    模板列(标记占位符的那个cell)
     * @param row         新增的行
     * @param cellSize    每行的列数量（用来补列补足的情况）
     * @param rowDataList 每行的数据
     */
    private void addRow(XWPFTableCell flagCell, XWPFTableRow row, int cellSize, List<String> rowDataList) {
        for (int i = 0; i < cellSize; i++) {
            XWPFTableCell cell = row.getCell(i);
            cell = cell == null ? row.createCell() : row.getCell(i);
            if (i < rowDataList.size()) {
                PoiWordUtils.copyCellAndSetValue(flagCell, cell, rowDataList.get(i));
            } else {
                // 数据不满整行时，添加空列
                PoiWordUtils.copyCellAndSetValue(flagCell, cell, "");
            }
        }
    }

    /**
     * 根据模板cell  添加重复行
     * @param startCell 模板列的开始位置
     * @param endCell 模板列的结束位置
     * @param currentRow 创建的新行
     * @param repeatFlagRow 模板列所在的行
     * @param rowDataList 每行的数据
     */
    private void addRowRepeat(int startCell, int endCell, XWPFTableRow currentRow, XWPFTableRow repeatFlagRow, List<String> rowDataList) {
        int cellSize = repeatFlagRow.getTableCells().size();
        for (int i = 0; i < cellSize; i++) {
            XWPFTableCell cell = currentRow.getCell(i);
            cell = cell == null ? currentRow.createCell() : currentRow.getCell(i);
            int flagCellIndex = i % (endCell - startCell + 1);
            XWPFTableCell repeatFlagCell = repeatFlagRow.getCell(flagCellIndex);
            if (i < rowDataList.size()) {
                PoiWordUtils.copyCellAndSetValue(repeatFlagCell, cell, rowDataList.get(i));
            } else {
                // 数据不满整行时，添加空列
                PoiWordUtils.copyCellAndSetValue(repeatFlagCell, cell, "");
            }
        }
    }

    /**
     * 复制段落
     *
     * @param sourcePar 原段落
     * @param targetPar
     * @param texts
     */
    private void copyParagraph(XWPFParagraph sourcePar, XWPFParagraph targetPar, String... texts) {

        targetPar.setAlignment(sourcePar.getAlignment());
        targetPar.setVerticalAlignment(sourcePar.getVerticalAlignment());

        // 设置布局
        targetPar.setAlignment(sourcePar.getAlignment());
        targetPar.setVerticalAlignment(sourcePar.getVerticalAlignment());

        if (texts != null && texts.length > 0) {
            String[] arr = texts;
            XWPFRun xwpfRun = sourcePar.getRuns().size() > 0 ? sourcePar.getRuns().get(0) : null;

            for (int i = 0, len = texts.length; i < len; i++) {
                String text = arr[i];
                XWPFRun run = targetPar.createRun();

                run.setText(text);

                run.setFontFamily(xwpfRun.getFontFamily());
                int fontSize = xwpfRun.getFontSize();
                run.setFontSize((fontSize == -1) ? DEFAULT_FONT_SIZE : fontSize);
                run.setBold(xwpfRun.isBold());
                run.setItalic(xwpfRun.isItalic());
            }
        }
    }

    public void setParamMap(Map<String, Object> paramMap) {
        this.paramMap = paramMap;
    }
}
```
## poi工具类PoiWordUtils

```
package per.qiao.utils.hutool.poi;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Assert;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;

import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * @author: qyp
 * Date: 2019-10-26 2:12
 */
public class PoiWordUtils {

    /**
     * 占位符第一个字符
     */
    public static final String PREFIX_FIRST = "$";

    /**
     * 占位符第二个字符
     */
    public static final String PREFIX_SECOND = "{";

    /**
     * 占位符的前缀
     */
    public static final String PLACEHOLDER_PREFIX = PREFIX_FIRST + PREFIX_SECOND;

    /**
     * 占位符后缀
     */
    public static final String PLACEHOLDER_END = "}";

    /**
     * 表格中需要动态添加行的独特标记
     */
    public static final String addRowText = "tbAddRow:";

    public static final String addRowRepeatText = "tbAddRowRepeat:";

    /**
     * 表格中占位符的开头 ${tbAddRow:  例如${tbAddRow:tb1}
     */
    public static final String addRowFlag = PLACEHOLDER_PREFIX + addRowText;

    /**
     * 表格中占位符的开头 ${tbAddRowRepeat:  例如 ${tbAddRowRepeat:0,2,0,1} 第0行到第2行，第0列到第1列 为模板样式
     */
    public static final String addRowRepeatFlag = PLACEHOLDER_PREFIX + addRowRepeatText;

    /**
     * 重复矩阵的分隔符  比如：${tbAddRowRepeat:0,2,0,1} 分隔符为 ,
     */
    public static final String tbRepeatMatrixSeparator = ",";

    /**
     * 占位符的后缀
     */
    public static final String PLACEHOLDER_SUFFIX = "}";

    /**
     * 图片占位符的前缀
     */
    public static final String PICTURE_PREFIX = PLACEHOLDER_PREFIX + "image:";

    /**
     * 判断当前行是不是标志表格中需要添加行
     *
     * @param row
     * @return
     */
    public static boolean isAddRow(XWPFTableRow row) {
        return isDynRow(row, addRowFlag);
    }

    /**
     * 添加重复模板动态行(以多行为模板)
     * @param row
     * @return
     */
    public static boolean isAddRowRepeat(XWPFTableRow row) {
        return isDynRow(row, addRowRepeatFlag);
    }

    private static boolean isDynRow(XWPFTableRow row, String dynFlag) {
        if (row == null) {
            return false;
        }
        List<XWPFTableCell> tableCells = row.getTableCells();
        if (tableCells != null) {
            XWPFTableCell cell = tableCells.get(0);
            if (cell != null) {
                String text = cell.getText();
                if (text != null && text.startsWith(dynFlag)) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 从参数map中获取占位符对应的值
     *
     * @param paramMap
     * @param key
     * @return
     */
    public static Object getValueByPlaceholder(Map<String, Object> paramMap, String key) {
        if (paramMap != null) {
            if (key != null) {
                return paramMap.get(getKeyFromPlaceholder(key));
            }
        }
        return null;
    }

    /**
     * 后去占位符的重复行列矩阵
     * @param key 占位符
     * @return {0,2,0,1}
     */
    public static String getTbRepeatMatrix(String key) {
        Assert.assertNotNull("占位符为空", key);
        String $1 = key.replaceAll("\\" + PREFIX_FIRST + "\\" + PREFIX_SECOND + addRowRepeatText + "(.*)" + "\\" + PLACEHOLDER_SUFFIX, "$1");
        return $1;
    }

    /**
     * 从占位符中获取key
     *
     * @return
     */
    public static String getKeyFromPlaceholder(String placeholder) {
        return Optional.ofNullable(placeholder).map(p -> p.replaceAll("[\\$\\{\\}]", "")).get();
    }

    public static void main(String[] args) {
        String s = "${aa}";
        s = s.replaceAll(PLACEHOLDER_PREFIX + PLACEHOLDER_SUFFIX , "");
        System.out.println(s);
//        String keyFromPlaceholder = getKeyFromPlaceholder("${tbAddRow:tb1}");
//        System.out.println(keyFromPlaceholder);
    }

    /**
     * 复制列的样式，并且设置值
     * @param sourceCell
     * @param targetCell
     * @param text
     */
    public static void copyCellAndSetValue(XWPFTableCell sourceCell, XWPFTableCell targetCell, String text) {
        //段落属性
        List<XWPFParagraph> sourceCellParagraphs = sourceCell.getParagraphs();
        if (sourceCellParagraphs == null || sourceCellParagraphs.size() <= 0) {
            return;
        }
        XWPFParagraph sourcePar = sourceCellParagraphs.get(0);
        XWPFParagraph targetPar = targetCell.getParagraphs().get(0);

        // 设置段落的样式
        targetPar.getCTP().setPPr(sourcePar.getCTP().getPPr());
//        CTTcBorders tcBorders = sourceCell.getCTTc().getTcPr().getTcBorders();
//        targetCell.getCTTc().getTcPr().setTcBorders(tcBorders);

        List<XWPFRun> sourceParRuns = sourcePar.getRuns();
        if (sourceParRuns != null && sourceParRuns.size() > 0) {
            // 如果当前cell中有run
            List<XWPFRun> runs = targetPar.getRuns();
            Optional.ofNullable(runs).ifPresent(rs -> rs.stream().forEach(r -> r.setText("", 0)));
            if (runs != null && runs.size() > 0) {
                runs.get(0).setText(text, 0);
            } else {
                XWPFRun cellR = targetPar.createRun();
                cellR.setText(text, 0);
                // 设置列的样式位模板的样式
                targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            }
            setTypeface(sourcePar, targetPar);
        } else {
            targetCell.setText(text);
        }
    }
    /**
     * 复制字体
     */
    private static void setTypeface(XWPFParagraph sourcePar, XWPFParagraph targetPar) {
        XWPFRun sourceRun = sourcePar.getRuns().get(0);
        String fontFamily = sourceRun.getFontFamily();
        //int fontSize = sourceRun.getFontSize();
        String color = sourceRun.getColor();
//        String fontName = sourceRun.getFontName();
        boolean bold = sourceRun.isBold();
        boolean italic = sourceRun.isItalic();
        int kerning = sourceRun.getKerning();
//        String style = sourcePar.getStyle();
        UnderlinePatterns underline = sourceRun.getUnderline();

        XWPFRun targetRun = targetPar.getRuns().get(0);
        targetRun.setFontFamily(fontFamily);
//        targetRun.setFontSize(fontSize == -1 ? 10 : fontSize);
        targetRun.setBold(bold);
        targetRun.setColor(color);
        targetRun.setItalic(italic);
        targetRun.setKerning(kerning);
        targetRun.setUnderline(underline);
        //targetRun.setFontSize(fontSize);
    }
    /**
     * 判断文本中时候包含$
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        if(text.indexOf(PLACEHOLDER_PREFIX)!= -1){
            check = true;
        }
        return check;
    }

    /**
     * 获得占位符替换的正则表达式
     * @return
     */
    public static String getPlaceholderReg(String text) {
        return "\\" + PREFIX_FIRST + "\\" + PREFIX_SECOND + text + "\\" + PLACEHOLDER_SUFFIX;
    }

    public static String getDocKey(String mapKey) {
        return PLACEHOLDER_PREFIX + mapKey + PLACEHOLDER_SUFFIX;
    }

    /**
     * 判断当前占位符是不是一个图片占位符
     * @param text
     * @return
     */
    public static boolean isPicture(String text) {
        return text.startsWith(PICTURE_PREFIX);
    }

    /**
     * 删除一行的列
     * @param row
     */
    public static void removeCells(XWPFTableRow row) {
        int size = row.getTableCells().size();
        try {
            for (int i = 0; i < size; i++) {
                row.removeCell(i);
            }
        } catch (Exception e) {

        }
    }
}
```
## 对图片的支持ImageEntity

```
package per.qiao.utils.hutool.poi;

import static per.qiao.utils.hutool.poi.ImageUtils.ImageType.PNG;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * 图片实体对象
 *
 * @author: qyp
 * Date: 2019-10-26 21:52
 */
public class ImageEntity {

    /**
     * 图片宽度
     */
    private int width = 400;

    /**
     * 图片高度
     */
    private int height = 300;

    /**
     * 图片地址
     */
    private String url;

    /**
     * 图片类型
     * @see ImageUtils.ImageType
     */
    private ImageUtils.ImageType typeId = PNG;

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public ImageUtils.ImageType getTypeId() {
        return typeId;
    }

    public void setTypeId(ImageUtils.ImageType typeId) {
        this.typeId = typeId;
    }
}
```
## 图片工具类ImageUtils

```
package per.qiao.utils.hutool.poi;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * @author: qyp
 * Date: 2019-10-26 21:55
 */
public class ImageUtils {

    /**
     * 图片类型枚举
     */
    enum ImageType {

        /**
         * 支持四种类型 JPG/JPEG, GIT, BMP, PNG
         */
        JPG("JPG", XWPFDocument.PICTURE_TYPE_JPEG),
        JPEG("JPEG", XWPFDocument.PICTURE_TYPE_JPEG),
        GIF("GIF", XWPFDocument.PICTURE_TYPE_GIF),
        BMP("BMP", XWPFDocument.PICTURE_TYPE_GIF),
        PNG("PNG", XWPFDocument.PICTURE_TYPE_PNG)
        ;
        private String name;
        private Integer typeId;
        ImageType(String name, Integer type) {
            this.name = name;
            this.typeId = type;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getTypeId() {
            return typeId;
        }

        public void setTypeId(Integer typeId) {
            this.typeId = typeId;
        }
    }


    public static void createPicture(XWPFRun run, String blipId, int id, int width, int height) {
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        CTInline inline = run.getCTR().addNewDrawing().addNewInline();

        String picXml = "" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "         </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";

        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch(XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }

}
```
## 测试类

```
package per.qiao.utils.hutool.poi;

import org.junit.Test;

import java.util.*;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * @author: qyp
 * Date: 2019-10-26 17:34
 */
public class DynWordUtilsTest {

    /**
     * 说明 普通占位符位${field}格式
     * 表格中的占位符为${tbAddRow:tb1}  tb1为唯一标识符
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) {

        // 模板全的路径
        String templatePaht = "E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\审查报告模板1023体检表.docx";
        // 输出位置
        String outPath = "e:\\22.docx";

        Map<String, Object> paramMap = new HashMap<>(16);
        // 普通的占位符示例 参数数据结构 {str,str}
        paramMap.put("title", "德玛西亚");
        paramMap.put("startYear", "2010");
        paramMap.put("endYear", "2020");
        paramMap.put("currentYear", "2019");
        paramMap.put("currentMonth", "10");
        paramMap.put("currentDate", "26");
        paramMap.put("name", "黑色玫瑰");

        // 段落中的动态段示例 [str], 支持动态行中添加图片
        List<Object> list1 = new ArrayList<>(Arrays.asList("2、list1_11111", "3、list1_2222", "${image:image0}"));
        ImageEntity imgEntity = new ImageEntity();
        imgEntity.setHeight(200);
        imgEntity.setWidth(300);
        imgEntity.setUrl("E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\image1.jpg");
        imgEntity.setTypeId(ImageUtils.ImageType.JPG);

        paramMap.put("image:image0", imgEntity);
        paramMap.put("list1", list1);

        List<String> list2 = new ArrayList<>(Arrays.asList("2、list2_11111", "3、list2_2222"));
        paramMap.put("list2", list2);

        // 表格中的参数示例 参数数据结构 [[str]]
        List<List<String>> tbRow1 = new ArrayList<>();
        List<String> tbRow1_row1 = new ArrayList<>(Arrays.asList("1、模块一", "分类1"));
        List<String> tbRow1_row2 = new ArrayList<>(Arrays.asList("2、模块二", "分类2"));
        tbRow1.add(tbRow1_row1);
        tbRow1.add(tbRow1_row2);
        paramMap.put(PoiWordUtils.addRowText + "tb1", tbRow1);

        List<List<String>> tbRow2 = new ArrayList<>();
        List<String> tbRow2_row1 = new ArrayList<>(Arrays.asList("指标c", "指标c的意见"));
        List<String> tbRow2_row2 = new ArrayList<>(Arrays.asList("指标d", "指标d的意见"));
        tbRow2.add(tbRow2_row1);
        tbRow2.add(tbRow2_row2);
        paramMap.put(PoiWordUtils.addRowText + "tb2", tbRow2);

        List<List<String>> tbRow3 = new ArrayList<>();
        List<String> tbRow3_row1 = new ArrayList<>(Arrays.asList("3", "耕地估值"));
        List<String> tbRow3_row2 = new ArrayList<>(Arrays.asList("4", "耕地归属", "平方公里"));
        tbRow3.add(tbRow3_row1);
        tbRow3.add(tbRow3_row2);
        paramMap.put(PoiWordUtils.addRowText + "tb3", tbRow3);

        // 支持在表格中动态添加图片
        List<List<String>> tbRow4 = new ArrayList<>();
        List<String> tbRow4_row1 = new ArrayList<>(Arrays.asList("03", "旅游用地", "18.8m2"));
        List<String> tbRow4_row2 = new ArrayList<>(Arrays.asList("04", "建筑用地"));
        List<String> tbRow4_row3 = new ArrayList<>(Arrays.asList("04", "${image:image3}"));
        tbRow4.add(tbRow4_row3);
        tbRow4.add(tbRow4_row1);
        tbRow4.add(tbRow4_row2);

        // 支持在表格中添加重复模板的行
        List<List<String>> tbRow5 = new ArrayList<>();
        List<String> tbRow5_row1 = new ArrayList<>(Arrays.asList("欢乐喜剧人"));
        List<String> tbRow5_row2 = new ArrayList<>(Arrays.asList("常远", "艾伦"));
        List<String> tbRow5_row3 = new ArrayList<>(Arrays.asList("岳云鹏", "孙越"));

        List<String> tbRow5_row4 = new ArrayList<>(Arrays.asList("诺克萨斯"));
        List<String> tbRow5_row5 = new ArrayList<>(Arrays.asList("德莱文", "诺手"));
        List<String> tbRow5_row6 = new ArrayList<>(Arrays.asList("男枪", "卡特琳娜"));

        tbRow5.add(tbRow5_row1);
        tbRow5.add(tbRow5_row2);
        tbRow5.add(tbRow5_row3);
        tbRow5.add(tbRow5_row4);
        tbRow5.add(tbRow5_row5);
        tbRow5.add(tbRow5_row6);
        paramMap.put("tbAddRowRepeat:0,2,0,1", tbRow5);

        ImageEntity imgEntity3 = new ImageEntity();
        imgEntity3.setHeight(100);
        imgEntity3.setWidth(100);
        imgEntity3.setUrl("E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\image1.jpg");
        imgEntity3.setTypeId(ImageUtils.ImageType.JPG);

        paramMap.put(PoiWordUtils.addRowText + "tb4", tbRow4);
        paramMap.put("image:image3", imgEntity3);

        // 图片占位符示例 ${image:imageid} 比如 ${image:image1}, ImageEntity中的值就为image:image1
        // 段落中的图片
        ImageEntity imgEntity1 = new ImageEntity();
        imgEntity1.setHeight(500);
        imgEntity1.setWidth(400);
        imgEntity1.setUrl("E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\image1.jpg");
        imgEntity1.setTypeId(ImageUtils.ImageType.JPG);
        paramMap.put("image:image1", imgEntity1);

        // 表格中的图片
        ImageEntity imgEntity2 = new ImageEntity();
        imgEntity2.setHeight(200);
        imgEntity2.setWidth(100);
        imgEntity2.setUrl("E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\image1.jpg");
        imgEntity2.setTypeId(ImageUtils.ImageType.JPG);
        paramMap.put("image:image2", imgEntity2);

        DynWordUtils.process(paramMap, templatePaht, outPath);
    }


    @Test
    public void testImage() {

        Map<String, Object> paramMap = new HashMap<>(16);
        String templatePaht = "E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\11.docx";
        String outPath = "e:\\3.docx";
        ImageEntity imgEntity1 = new ImageEntity();
        imgEntity1.setHeight(500);
        imgEntity1.setWidth(400);
        imgEntity1.setUrl("E:\\Java4IDEA\\comm_test\\commutil\\src\\main\\resources\\wordtemplate\\image1.jpg");
        imgEntity1.setTypeId(ImageUtils.ImageType.JPG);

        paramMap.put("image:img1", imgEntity1);
        DynWordUtils.process(paramMap, templatePaht, outPath);
    }
}
```
## 模板图
