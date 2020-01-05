package word;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * @author sparkchan
 * @date 2020/1/4
 */
public class WordUtils {
    public static String FIELD_NAME = "fieldName";
    public static String GENERAL_TEMPLATE = "#{fieldName}";
    public static String TABLE_TEMPLATE_START = "<iterator(fieldName)>";
    public static String TABLE_TEMPLATE_END = "</iterator(fieldName)>";

    public static XWPFDocument createWord(String filePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(filePath);
        return createWord(inputStream);
    }

    public static XWPFDocument createWord(InputStream inputStream) throws IOException {
        XWPFDocument document = new XWPFDocument(inputStream);
        return document;
    }

    public static void fillTemplates(WordTemplateReplaceable wordTemplateEntity, XWPFDocument document) {
        // 参数校验
        Objects.requireNonNull(document, "模版对象不能为空!");
        Objects.requireNonNull(wordTemplateEntity, "实体对象不能为空!");
        requireReplaceable(wordTemplateEntity);
        // 填充一般字符串模版
        fillGeneralTemplates(wordTemplateEntity, document);
        // 填充表格模板
        fillTableTemplates(wordTemplateEntity, document);
    }

    private static void requireReplaceable(Object obj) {
        if (obj instanceof WordTemplateReplaceable) {
            return;
        }
        throw new NotWordTemplateReplaceableException("该实体对象未实现 WordTNotWordTemplateReplaceableExceptionemplateReplaceale 接口, 不能用于填充模版！");
    }

    private static void fillGeneralTemplates(WordTemplateReplaceable wordTemplateEntity, XWPFDocument document) {
        List<Field> fields = Arrays.asList(wordTemplateEntity.getClass().getDeclaredFields());
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (Field field : fields) {
            for (XWPFParagraph paragraph : paragraphs) {
                replaceGeneralTemplate(wordTemplateEntity, field, paragraph);
            }
        }
    }

    private static void fillTableTemplates(WordTemplateReplaceable wordTemplateEntity, XWPFDocument document) {
        List<Field> fields = new ArrayList<>(Arrays.asList(wordTemplateEntity.getClass().getDeclaredFields()));
        Iterator<Field> iterator = fields.iterator();
        while (iterator.hasNext()) {
            Field field = iterator.next();
            if (Iterable.class.isAssignableFrom(field.getType())) {
                continue;
            }
            iterator.remove();
        }
        List<XWPFTable> tables = document.getTables();
        for (Field field : fields) {
            for (XWPFTable table : tables) {
                replaceTableTemplate(wordTemplateEntity, field, table);
            }
        }
    }

    private static void replaceTableTemplate(Object fieldDeclaredObj, Field field, XWPFTable table) {
        String fieldName = field.getName();
        String tableTemplateStart = TABLE_TEMPLATE_START.replace(FIELD_NAME, fieldName);
        String tableTemplateEnd = TABLE_TEMPLATE_END.replace(FIELD_NAME, fieldName);
        String text = table.getText().replace(" ", "");
        boolean isReplaceableTable = text.contains(tableTemplateStart) && text.contains(tableTemplateEnd);
        if (isReplaceableTable) {
            Map<String, Integer> name2IndexMap = getTemplateName2IndexMap(table);
            clearTableRow(table);
            try {
                field.setAccessible(true);
                Iterable<? extends WordTemplateReplaceable> iterable = (Iterable<? extends WordTemplateReplaceable>) field
                        .get(fieldDeclaredObj);
                boolean isFirstRow = true;
                for (WordTemplateReplaceable wordTemplate : iterable) {
                    XWPFTableRow row = table.createRow();
                    List<Field> fields = Arrays.asList(wordTemplate.getClass().getDeclaredFields());
                    if (isFirstRow) {
                        for (int i = 0; i < fields.size(); i++) {
                            row.createCell();
                        }
                        isFirstRow = false;
                    }
                    for (Field subField : fields) {
                        Integer index = name2IndexMap.get(GENERAL_TEMPLATE.replace(FIELD_NAME, subField.getName()));
                        XWPFTableCell cell = row.getCell(index);
                        subField.setAccessible(true);
                        String content = (String) subField.get(wordTemplate);
                        cell.setText(content);
                    }
                }
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    private static Map<String, Integer> getTemplateName2IndexMap(XWPFTable table) {
        Map<String, Integer> name2IndexMap = new HashMap<>(16);
        XWPFTableRow row = table.getRow(1);
        List<XWPFTableCell> cells = row.getTableCells();
        for (int i = 0; i < cells.size(); i++) {
            name2IndexMap.put(cells.get(i).getText(), i);
        }
        return name2IndexMap;
    }

    private static void clearTableRow(XWPFTable table) {
        while (table.getRows().size() > 0) {
            table.removeRow(0);
        }
    }

    private static void replaceGeneralTemplate(Object fieldDeclaredObj, Field field, XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        String fieldName = field.getName();
        String template = GENERAL_TEMPLATE.replace(FIELD_NAME, fieldName);
        for (XWPFRun run : runs) {
            // TODO 这里默认获取第一个 Text, 看以后能不能改进
            String text = run.getText(0);
            text = text.replace(" ", "");
            if (text.contains(template)) {
                field.setAccessible(true);
                try {
                    text = text.replace(template, (String) field.get(fieldDeclaredObj));
                    run.setText(text, 0);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    // TODO 增加图片

    // TODO 增加表格

    // TODO 增加段落
}
