package com.yyz.util.wordutil;

import com.yyz.util.constant.Constants;
import org.apache.poi.xwpf.usermodel.*;

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
     * 判断当前行是不是标志表格中需要添加行
     *
     * @param row
     * @return
     */
    public static boolean isAddRow(XWPFTableRow row) {
        if (row != null) {
            List<XWPFTableCell> tableCells = row.getTableCells();
            if (tableCells != null) {
                for (int i = 0; i < tableCells.size(); i++) {
                    XWPFTableCell cell = tableCells.get(i);
                    if (cell != null) {
                        // 表格里面包含表格特殊处理,也是动态添加行
                        List<IBodyElement> bodyElements = cell.getBodyElements();
                        // 表格里的cell只有一个元素,自己判断是否是动态行即可
                        if (bodyElements.size() == 1) {
                            String text = cell.getText();
                            if (text != null && text.startsWith(Constants.ADD_ROW_FLAG)) {
                                return true;
                            }
                        }else {
                            // 表格里有多个元素,需要遍历判断是否是动态行
                            for (int j = 0; j < bodyElements.size(); j++) {
                                IBodyElement bodyElement = bodyElements.get(j);
                                if (bodyElement instanceof XWPFTable) {
                                    XWPFTable xwpfTable =(XWPFTable)bodyElement;
                                    List<XWPFTableRow> rows = xwpfTable.getRows();
                                    // 此处写死,直接获取表格中第二行的数据,即默认有标题行,潘森是否是动态添加行
                                    String text = rows.get(1).getCell(0).getText();
                                    if (text != null && text.startsWith(Constants.ADD_ROW_FLAG)) {
                                        return true;
                                    }
                                }
                            }
                        }
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
        if(text.indexOf(Constants.PLACEHOLDER_PREFIX)!= -1){
            check = true;
        }
        return check;
    }

    /**
     * 获得占位符替换的正则表达式
     * @return
     */
    public static String getPlaceholderReg(String text) {
        return "\\" + Constants.PREFIX_FIRST + "\\" + Constants.PREFIX_SECOND + text + "\\" + Constants.PLACEHOLDER_SUFFIX;
    }

    public static String getDocKey(String mapKey) {
        return Constants.PLACEHOLDER_PREFIX + mapKey + Constants.PLACEHOLDER_SUFFIX;
    }

    /**
     * 判断当前占位符是不是一个图片占位符
     * @param text
     * @return
     */
    public static boolean isPicture(String text) {
        return text.startsWith(Constants.PICTURE_PREFIX);
    }
}
