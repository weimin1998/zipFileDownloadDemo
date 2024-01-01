package com.example.zipfileDownloadDemo;

import org.apache.poi.xwpf.usermodel.*;

import java.util.Map;

public class TemplateUtils {

    /**
     * 填充段落中的占位符
     * @param paragraph
     * @param dataMap
     */
    public static void replacePlaceholder(XWPFParagraph paragraph, Map<String, String> dataMap) {
        if (paragraph.getText().contains("$")) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.text();
                for (String key : dataMap.keySet()) {
                    String placeHolder = "${" + key + "}";
                    if (text.contains(placeHolder)) {
                        text = text.replace(placeHolder, dataMap.get(key));
                        run.setText(text, 0);
                    }
                }
            }
        }
    }

    public static void replacePlaceholder(XWPFDocument xwpfDocument, Map<String, String> dataMap) {
        for (XWPFParagraph paragraph : xwpfDocument.getParagraphs()) {
            replacePlaceholder(paragraph, dataMap);
        }
    }

    /**
     * 填充表格中的占位符
     * @param table
     * @param dataMap
     */
    public static void replacePlaceholderInTable(XWPFTable table, Map<String, String> dataMap) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell tableCell : row.getTableCells()) {
                for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                    replacePlaceholder(paragraph, dataMap);
                }
            }
        }
    }

}
