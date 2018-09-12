package Utils;

import java.math.BigInteger;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class TableUtils {
	/**
	 * @param document
	 * @param rowSize
	 * @param cellSize
	 * @param isSetColWidth
	 * @param colWidths
	 * @return
	 */
	public static XWPFTable createTable(XWPFDocument document, int rowSize, int cellSize,
			int[] colWidths) {
		XWPFTable table = document.createTable(rowSize, cellSize);
		if (ArrayUtils.isNotEmpty(colWidths)) {
			int totalWidth = 0;
			CTTbl ttbl = table.getCTTbl();
			CTTblGrid tblGrid = ttbl.addNewTblGrid();
			for (int j = 0, len = Math.min(cellSize, colWidths.length); j < len; j++) {
				CTTblGridCol gridCol = tblGrid.addNewGridCol();
				gridCol.setW(new BigInteger(String.valueOf(colWidths[j])));
				totalWidth += colWidths[j];
			}
			table.setWidth(totalWidth);
		}
		return table;
	}

	/**
	 * @param cell
	 * @param width
	 * @param sttblWidth
	 */
	public static void setCellWidth(XWPFTableCell cell, int width, STTblWidth.Enum sttblWidth) {
		CTTcPr cttcPr = getCellCTTcPr(cell);
		CTTblWidth tblWidth = cttcPr.isSetTcW() ? cttcPr.getTcW() : cttcPr.addNewTcW();
		if(width>0) {
			tblWidth.setW(BigInteger.valueOf(width));
		}
		if (sttblWidth != null) {
			tblWidth.setType(sttblWidth);
		} else {
			tblWidth.setType(STTblWidth.DXA);
		}

	}

	/**
	 * @param cell
	 * @param stVerticalJc
	 */
	public static void setCellVAlign(XWPFTableCell cell, STVerticalJc.Enum stVerticalJc) {
		CTTcPr cttcPr = getCellCTTcPr(cell);
		CTVerticalJc vJc = cttcPr.isSetVAlign() ? cttcPr.getVAlign() : cttcPr.addNewVAlign();
		if (stVerticalJc != null) {
			vJc.setVal(stVerticalJc);
		} else {
			vJc.setVal(STVerticalJc.CENTER);
		}
	}

	/**
	 * @param cell
	 * @param content
	 * @param stJc 對其方式
	 */
	public static void addCellContent(XWPFTableCell cell, String content, ParagraphAlignment pAlignment, STUnderline.Enum stUnderline, int fontSize) {
		XWPFParagraph p = cell.addParagraph();
		XWPFRun run = p.createRun();
		ParagraphUtils.setParagraphTextStyleInfo(p, run, null, false, true, null, stUnderline, false, false, false, false, false, false, false, false, false, null, false, null, false, null, content, null, 0, 0, 0, fontSize);
		
		if(pAlignment != null) {
			p.setAlignment(pAlignment);
		}
		run.setText(content);
//		CTTc cttc = cell.getCTTc();
//		CTP p = cttc.addNewP();
//		CTR r = p.addNewR();
//		CTText text = r.addNewT();
//		text.setStringValue(content);
//		if(stJc != null) {
//			CTPPr ppr = getCTPCTPPr(p);
//			CTJc jc = getCTPPrCTJc(ppr);
//			jc.setVal(stJc);
//		}
//		
//		
//		if(stUnderline != null) {
//			CTRPr rpr = getCTRCTRPr(r);
//			rpr.addNewU().setVal(stUnderline);
//		}
	}
	
	/**
	 * 在表格內加入摘要號碼
	 * @param cell
	 * @param numId
	 * @param ilvl
	 */
	public static void addNumberingToCell(XWPFTableCell cell, BigInteger numId, BigInteger ilvl) {
		CTTc cttc = cell.getCTTc();
		CTP p = cttc.getPArray(0);
		CTPPr ppr = getCTPCTPPr(p);
		CTNumPr numPr = ParagraphUtils.getCTPPrCTNumPr(ppr);
		ParagraphUtils.setNumPr(numPr, numId, ilvl);
	}

	/**
	 * @param table
	 * @param row
	 * @param fromCell
	 * @param toCell
	 */
	public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCell) {
				// The first merged cell is set with RESTART merge value
				getCellCTTcPr(cell).addNewHMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one,are set with CONTINUE
				getCellCTTcPr(cell).addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * @param table
	 * @param col
	 * @param fromRow
	 * @param toRow
	 */
	public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) {
				// The first merged cell is set with RESTART merge value
				getCellCTTcPr(cell).addNewVMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one,are set with CONTINUE
				getCellCTTcPr(cell).addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * @param table
	 * @param index
	 */
	public static void deleteTableRow(XWPFTable table, int index) {
		table.removeRow(index);
	}

	/**
	 * @param table
	 * @param rowIndex
	 * @param col
	 * @return
	 */
	public static XWPFTableCell getCell(XWPFTable table, int rowIndex, int col) {
		return table.getRow(rowIndex).getCell(col);
	}

	/**
	 * @param row
	 * @param hight
	 * @param heigthEnum
	 */
	public static void setRowHeight(XWPFTableRow row, String hight, STHeightRule.Enum heigthEnum) {
		CTTrPr trPr = getRowCTTrPr(row);
		CTHeight trHeight;
		if (trPr.getTrHeightList() != null && trPr.getTrHeightList().size() > 0) {
			trHeight = trPr.getTrHeightList().get(0);
		} else {
			trHeight = trPr.addNewTrHeight();
		}
		trHeight.setVal(new BigInteger(hight));
		if (heigthEnum != null) {
			trHeight.setHRule(heigthEnum);
		}
	}

	/**
	 * @param table
	 * @param left
	 * @param top
	 * @param right
	 * @param bottom
	 */
	public static void setTableBorders(XWPFTable table, CTBorder left, CTBorder top, CTBorder right, CTBorder bottom) {
		CTTblBorders tblBorders = getTableBorders(table);
		if (left != null) {
			tblBorders.setLeft(left);
		}
		if (top != null) {
			tblBorders.setTop(top);
		}
		if (right != null) {
			tblBorders.setRight(right);
		}
		if (bottom != null) {
			tblBorders.setBottom(bottom);
		}
	}
	
	
	
	/**
	 * @param table
	 */
	public static void unsetTableBorders(XWPFTable table) {
		table.getCTTbl().getTblPr().unsetTblBorders();
	}

	public static CTNumPr getCTPPrCTNumPr(CTPPr ppr) {
		return ppr.isSetNumPr()? ppr.getNumPr() : ppr.addNewNumPr();
	}
	
	public static CTJc getCTPPrCTJc(CTPPr ppr) {
		return ppr.isSetJc() ? ppr.getJc() : ppr.addNewJc();
	}

	public static CTPPr getCTPCTPPr(CTP p) {
		return p.isSetPPr() ? p.getPPr() : p.addNewPPr();
	}

	public static CTRPr getCTRCTRPr(CTR r) {
		return r.isSetRPr()? r.getRPr():r.addNewRPr();
	}

	public static CTTcPr getCellCTTcPr(XWPFTableCell cell) {
		CTTc cttc = cell.getCTTc();
		CTTcPr tcPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
		return tcPr;
	}

	public static CTTrPr getRowCTTrPr(XWPFTableRow row) {
		CTRow ctRow = row.getCtRow();
		CTTrPr trPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
		return trPr;
	}

	public static CTTblBorders getTableBorders(XWPFTable table) {
		CTTblPr tblPr = getTableCTTblPr(table);
		CTTblBorders tblBorders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
		return tblBorders;
	}

	public static CTTblPr getTableCTTblPr(XWPFTable table) {
		CTTbl ttbl = table.getCTTbl();
		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
		return tblPr;
	}
}
