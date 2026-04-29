package ru.runa.wfe.office.doc;

import com.google.common.collect.Lists;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import ru.runa.wfe.office.OfficeProperties;
import ru.runa.wfe.var.MapDelegableVariableProvider;
import ru.runa.wfe.var.VariableProvider;
import ru.runa.wfe.var.dto.WfVariable;

public class DocxFileChanger {
    private final DocxConfig config;
    private final MapDelegableVariableProvider variableProvider;
    private final XWPFDocument document;

    public DocxFileChanger(DocxConfig config, VariableProvider variableProvider, InputStream templateInputStream) throws IOException {
        this.config = config;
        this.variableProvider = new MapDelegableVariableProvider(new HashMap<String, Object>(), variableProvider);
        this.document = new XWPFDocument(templateInputStream);
    }

    public XWPFDocument changeAll() {
        for (XWPFHeader header : document.getHeaderList()) {
            changeBodyElements(header.getBodyElements());
        }
        changeBodyElements(document.getBodyElements());
        for (XWPFFooter footer : document.getFooterList()) {
            changeBodyElements(footer.getBodyElements());
        }
        return document;
    }

    private void changeBodyElements(List<IBodyElement> bodyElements) {
        List<XWPFParagraph> paragraphs = Lists.newArrayList();

        for (IBodyElement bodyElement : new ArrayList<IBodyElement>(bodyElements)) {
            if (bodyElement instanceof XWPFParagraph) {
                paragraphs.add((XWPFParagraph) bodyElement);
                continue;
            }

            if (!paragraphs.isEmpty()) {
                DocxUtils.replaceInParagraphs(config, variableProvider, paragraphs);
                paragraphs.clear();
            }

            if (bodyElement instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) bodyElement;
                DocxConfig.TableConfig tableConfig = findTableConfig(table);

                List<Integer> emptyRowIndices = Lists.newArrayList();
                List<XWPFTableRow> rows = table.getRows();

                for (int i = 0; i < rows.size(); i++) {
                    XWPFTableRow row = rows.get(i);
                    List<XWPFTableCell> cells = row.getTableCells();

                    TableExpansionOperation tableExpansionOperation = new TableExpansionOperation(row);
                    boolean templateRowIsEmpty = true;

                    for (int columnIndex = 0; columnIndex < cells.size(); columnIndex++) {
                        XWPFTableCell cell = cells.get(columnIndex);
                        String cellText = cell.getText();

                        ColumnExpansionOperation operation = DocxUtils.parseIterationOperation(config, variableProvider, cellText,
                                new ColumnExpansionOperation());

                        if (operation == null && tableConfig != null && tableConfig.getListVariableName() != null
                                && !tableConfig.getListVariableName().isEmpty()) {
                            String attributeName = "";
                            if (tableConfig.getColumns().size() > columnIndex) {
                                attributeName = tableConfig.getColumns().get(columnIndex);
                            }

                            ColumnExpansionOperation colOp = new ColumnExpansionOperation();
                            colOp.setIterateBy(IterateBy.items);
                            colOp.setContainerVariableName(tableConfig.getListVariableName());
                            colOp.setContainerSelector(attributeName);
                            WfVariable listVariable = variableProvider.getVariable(tableConfig.getListVariableName());
                            if (listVariable != null) {
                                colOp.setContainerVariable(listVariable);
                            }

                            operation = colOp;
                        }

                        if (operation != null && operation.isValid()) {
                            tableExpansionOperation.addOperation(columnIndex, operation);
                        }

                        String text0 = null;
                        if (tableExpansionOperation.getRows() > 0) {
                            text0 = tableExpansionOperation.getStringValue(config, variableProvider, columnIndex, 0);
                        }

                        if (templateRowIsEmpty && text0 != null && !text0.trim().isEmpty()) {
                            templateRowIsEmpty = false;
                        }

                        if (!java.util.Objects.equals(text0, cell.getText())) {
                            DocxUtils.setCellText(cell, text0 == null ? "" : text0);
                        }
                    }

                    if (tableExpansionOperation.getRows() == 0) {
                        for (XWPFTableCell cell : cells) {
                            DocxUtils.replaceInParagraphs(config, variableProvider, cell.getParagraphs());
                        }
                        continue;
                    }

                    int rowIndexSafe = templateRowIsEmpty ? 0 : 1;

                    for (int rowIndex = 1; rowIndex < tableExpansionOperation.getRows(); rowIndex++) {
                        boolean wholeRowIsEmpty = true;

                        XWPFTableRow dynamicRow = table.createRow();

                        int columns = Math.min(dynamicRow.getTableCells().size(), tableExpansionOperation.getTemplateRow().getTableCells().size());
                        for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                            String text;

                            if (tableExpansionOperation.isIndexOrNumberColumnExpansitonOperation(columnIndex)) {
                                text = tableExpansionOperation.getStringValue(config, variableProvider, columnIndex, rowIndexSafe);
                            } else {
                                text = tableExpansionOperation.getStringValue(config, variableProvider, columnIndex, rowIndex);
                            }

                            XWPFTableCell templateCell = columnIndex < tableExpansionOperation.getTemplateRow().getTableCells().size()
                                    ? tableExpansionOperation.getTemplateCell(columnIndex)
                                    : null;
                            DocxUtils.setCellText(dynamicRow.getCell(columnIndex), text, templateCell);

                            if (wholeRowIsEmpty && text != null && !text.trim().isEmpty()) {
                                wholeRowIsEmpty = false;
                            }
                        }

                        if (wholeRowIsEmpty) {
                            emptyRowIndices.add(table.getRows().indexOf(dynamicRow));
                        } else {
                            rowIndexSafe++;
                        }
                    }
                }

                for (Integer i : Lists.reverse(emptyRowIndices)) {
                    table.removeRow(i);
                }
            }
        }

        if (!paragraphs.isEmpty()) {
            DocxUtils.replaceInParagraphs(config, variableProvider, paragraphs);
            paragraphs.clear();
        }
    }

    private DocxConfig.TableConfig findTableConfig(XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                String text = cell.getText();
                for (Map.Entry<String, DocxConfig.TableConfig> entry : config.getTables().entrySet()) {
                    if (text.contains("${" + entry.getKey() + "}")) {
                        return entry.getValue();
                    }
                }
            }
        }
        return null;
    }
}