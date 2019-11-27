package per.qiao.utils.hutool.poi;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

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
     * 表格中需要动态添加行的独特标记
     */
    public static final String addRowText = "tbAddRow:";

    /**
     * 表格中占位符的开头 ${tbAddRow:  例如${tbAddRow:tb1}
     */
    public static final String addRowFlag = PLACEHOLDER_PREFIX + addRowText;

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
        if (row != null) {
            List<XWPFTableCell> tableCells = row.getTableCells();
            if (tableCells != null) {
                XWPFTableCell cell = tableCells.get(0);
                if (cell != null) {
                    String text = cell.getText();
                    if (text != null && text.startsWith(addRowFlag)) {
                        return true;
                    }
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
        } else {
            targetCell.setText(text);
        }
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
}
