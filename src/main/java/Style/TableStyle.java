package Style;

public class TableStyle {
	private String[][] content;
	private int rowNum;
	private int columnNum;
	private int tableWidth;

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public int getColumnNum() {
		return columnNum;
	}

	public void setColumnNum(int columnNum) {
		this.columnNum = columnNum;
	}

	public String[][] getContent() {
		return content;
	}

	public void setContent(String[][] content) {
		this.content = content;
	}

	public int getTableWidth() {
		return tableWidth;
	}

	public void setTableWidth(int tableWidth) {
		this.tableWidth = tableWidth;
	}
}
