package per.qiao.utils.hutool.poi;

import cn.hutool.core.util.ArrayUtil;
import cn.hutool.poi.word.Word07Writer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
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

    private Word07Writer writer = null;
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
        writer = new Word07Writer(inFile);
        templateDoc = writer.getDoc();

        try (FileOutputStream outStream = new FileOutputStream(outPath)) {
            parseTemplateWord();
            templateDoc.write(outStream);
        } catch (Exception e) {
            StackTraceElement[] stackTrace = e.getStackTrace();

            String className = stackTrace[0].getClassName();
            String methodName = stackTrace[0].getMethodName();
            String fileName = stackTrace[0].getFileName();
            int lineNumber = stackTrace[0].getLineNumber();

            logger.error("错误：第:{}行, 类名:{}, 方法名:{}, 字段名:{}", lineNumber, className, methodName, fileName);
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
        for (XWPFTableRow row : rows) {
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
            for (XWPFTableRow tableRow : xwpfTableRows) {
                delAndJudgeRow(table, paramMap, tableRow);
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
     * 添加行
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
     * 根据模板cell添加新行
     *
     * @param flagCell    模板列(标记占位符的那个cell)
     * @param row         新增的行
     * @param cellSize    每行的列数量（用来补列补足的情况）
     * @param rowDataList 每行的数据
     * @throws Exception
     */
    private void addRow(XWPFTableCell flagCell, XWPFTableRow row, int cellSize, List<String> rowDataList) throws Exception {
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

        if (ArrayUtil.isNotEmpty(texts)) {
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
