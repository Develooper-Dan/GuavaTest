package excel;

import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import net.objecthunter.exp4j.Expression;
import net.objecthunter.exp4j.ExpressionBuilder;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel3000 {
    Table<Integer, Integer, String> table;
    private final Pattern cellRefPattern = Pattern.compile("\\$([a-zA-Z]{1,2}\\d+)");
    private final Pattern expressionPattern = Pattern.compile(".*[^\\d\\+\\-\\*\\/\\^\\.\\s]+.*");

    public void createTable() {
        HashBasedTable<Integer, Integer, String> newTable = HashBasedTable.create();
        this.table = newTable;
    }

    private Integer getColumnFromString(String cell) {
        String colString = cell.replaceAll("\\d", "").toUpperCase();
        int codePointBeforeAbcStarts = "A".codePointAt(0);
        if(!colString.isEmpty() && !colString.isBlank() && colString.length() <= 2) {
            if (colString.length() == 1) {
                return colString.codePointAt(0) - codePointBeforeAbcStarts;
            } else if (colString.length() == 2){
                int valFirstLetter = (colString.codePointAt(0) - codePointBeforeAbcStarts);
                int valSecondLetter = colString.codePointAt(1) - codePointBeforeAbcStarts;
                return ((valFirstLetter + 1) * 26) + valSecondLetter; // ' +1 cause cols are 0-indexed
            }
        }
        System.out.println("Column identifier missing or out of bounds");
        return -1;
    }

    public void setCell(String cell, String content) {
        Integer row = Integer.parseInt(cell.replaceAll("\\D", ""));
        Integer column = getColumnFromString(cell);
        table.put(row, column, content);
    }

    public String getCellAt(int row, int column) {
            return table.get(row, column);
    }
    public String getCellAt(String cell) {
        Integer row = Integer.parseInt(cell.replaceAll("\\D", ""));
        Integer column = getColumnFromString(cell);
        return table.get(row, column);
    }

    public Table<Integer, Integer, String> evaluate() {
        Table<Integer, Integer, String> calculatedTable = HashBasedTable.create();
        for(Table.Cell<Integer, Integer, String> cell : table.cellSet()){
            String content = "";
            if(!cell.getValue().startsWith("=")) {
                content += cell.getValue();
            } else {
                Matcher matcher = cellRefPattern.matcher(cell.getValue().substring(1));
                String evaluableString = matcher.replaceAll(result -> getCellAt(result.group(1)));
                if (evaluableString.matches(expressionPattern.pattern())){
                    content += evaluableString;
                } else {
                    try{
                        Expression exp = new ExpressionBuilder(evaluableString).build();
                        content += String.valueOf(exp.evaluate());
                    } catch(IllegalArgumentException e){
                        content += "#WERT";
                    }
                }
            }
            calculatedTable.put(cell.getRowKey(), cell.getColumnKey(), content);
        }
        return calculatedTable;
    }
}
