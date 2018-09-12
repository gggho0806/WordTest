package WordTest.WordTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import Utils.NumbericListUtils;
import Utils.ParagraphUtils;
import Utils.TableUtils;

public class NestListedNumberingList {
	public static void main(String[] args) throws IOException {
		// Blank Document
		XWPFDocument document = new XWPFDocument();
		// Write the Document in file system
		FileOutputStream out = new FileOutputStream(new File("NestList.docx"));

		// 設定統一字體
		XWPFStyles styles = document.createStyles();
		CTFonts fonts = CTFonts.Factory.newInstance();
		fonts.setHAnsi("標楷體");
		fonts.setAscii("Times new roman");
		fonts.setEastAsia("標楷體");
		styles.setDefaultFonts(fonts);

		String[][] indexContent = { { "0", "項目", "頁次", "財務報告\n附註編號" }, { "0", "封 面", "1", "-" },
				{ "0", "目 錄", "2", "-" }, { "0", "會計師核閱報告", "3～ 4", "-" }, { "1", "公司沿革", "12", "一" },
				{ "1", "通過財務報告之日期及程序", "12", "一" }, { "1", "重大之期後事項", "-", "-" }, { "1", "重大之期後事項", "-", "-" },
				{ "1", "重大之期後事項", "-", "-" }, { "1", "重大之期後事項", "-", "-" }, { "1", "重大之期後事項", "-", "-" },
				{ "1", "重大之期後事項", "-", "-" }, { "1", "重大之期後事項", "-", "-" }, { "1", "重大之期後事項", "-", "-" },
				{ "1", "重大之期後事項", "-", "-" }, { "1", "重大之期後事項", "-", "-" },
				{ "0", "會計師核閱報告", "3～ 4", "-" }, { "1", "重大之期後事項", "-", "-" }};

		int[] colWidths = {3*1440,2*1440,2*1440};
		XWPFTable table = TableUtils.createTable(document, indexContent.length, indexContent[0].length-1, colWidths);
		TableUtils.unsetTableBorders(table);

		// 設定摘要編碼的格式
		String[] lvlTexts = {"{num}、", "({num})"};
		
		BigInteger numID = NumbericListUtils.addAbstractNum(document, STNumberFormat.CHINESE_COUNTING_THOUSAND, lvlTexts, 480, BigInteger.valueOf(0), BigInteger.valueOf(1));
		
		for (int i = 0; i < indexContent.length; i++) {
			if(i==0) {
				// 取得某一格
				XWPFTableCell cell = TableUtils.getCell(table, 0, 0);
				// 清除空白
				cell.removeParagraph(0);
				// 設定該格的寬度
				TableUtils.setCellWidth(cell, 3*1440, STTblWidth.DXA);
				// 設定垂直對準方式
				TableUtils.setCellVAlign(cell, STVerticalJc.BOTTOM);
				// 設定對其方式，要注意的是CTP相當於一個段落，而且每個Cell都預設有一個CTP，不同段落可以有不同對其方式
				// 設定內容、樣式和對其方式
				TableUtils.addCellContent(cell, indexContent[0][1], ParagraphAlignment.DISTRIBUTE, STUnderline.SINGLE, 15);
				
				cell = TableUtils.getCell(table, 0, 1);
				// 清除空白
				cell.removeParagraph(0);
				// 設定該格的寬度
				TableUtils.setCellWidth(cell, 2*1440, STTblWidth.DXA);
				// 設定垂直對準方式
				TableUtils.setCellVAlign(cell, STVerticalJc.BOTTOM);
				// 設定對其方式，要注意的是CTP相當於一個段落，而且每個Cell都預設有一個CTP，不同段落可以有不同對其方式
				// 設定內容、樣式和對其方式
				TableUtils.addCellContent(cell, indexContent[0][2], ParagraphAlignment.DISTRIBUTE, STUnderline.SINGLE, 15);
				
				String[] splitTemp = indexContent[0][3].split("\n");
				cell = TableUtils.getCell(table, 0, 2);
				// 清除空白
				cell.removeParagraph(0);
				// 設定該格的寬度
				TableUtils.setCellWidth(cell, 2*1440, STTblWidth.DXA);
				// 設定垂直對準方式
				TableUtils.setCellVAlign(cell, STVerticalJc.BOTTOM);
				// 設定對其方式，要注意的是CTP相當於一個段落，而且每個Cell都預設有一個CTP，不同段落可以有不同對其方式
				// 設定內容、樣式和對其方式
				TableUtils.addCellContent(cell, splitTemp[0], ParagraphAlignment.DISTRIBUTE, null, 15);
				TableUtils.addCellContent(cell, splitTemp[1], ParagraphAlignment.DISTRIBUTE, STUnderline.SINGLE, 15);
				
			}else {
				XWPFTableCell cell = TableUtils.getCell(table, i, 0);
				// 清除空白
				cell.removeParagraph(0);
				// 設定對其方式，要注意的是CTP相當於一個段落，而且每個Cell都預設有一個CTP，不同段落可以有不同對其方式
				// 設定內容、樣式和對其方式
				TableUtils.addCellContent(cell, indexContent[i][1], null, null, 10);
				
				int ilvl = Integer.valueOf(indexContent[i][0]);
				TableUtils.addNumberingToCell(cell, numID, BigInteger.valueOf(ilvl));
				
				cell = TableUtils.getCell(table, i, 1);
				// 清除空白
				cell.removeParagraph(0);
				// 設定對其方式，要注意的是CTP相當於一個段落，而且每個Cell都預設有一個CTP，不同段落可以有不同對其方式
				// 設定內容、樣式和對其方式
				TableUtils.addCellContent(cell, indexContent[i][2], ParagraphAlignment.CENTER, null, 10);
				
				cell = TableUtils.getCell(table, i, 2);
				// 清除空白
				cell.removeParagraph(0);
				// 設定對其方式，要注意的是CTP相當於一個段落，而且每個Cell都預設有一個CTP，不同段落可以有不同對其方式
				// 設定內容、樣式和對其方式
				TableUtils.addCellContent(cell, indexContent[i][3], ParagraphAlignment.CENTER, null, 10);
			}
		}

		// 設定摘要編碼的格式
		
		// 設定第一階層
		BigInteger numIDOne = NumbericListUtils.addAbstractNum(document, STNumberFormat.CHINESE_COUNTING_THOUSAND, lvlTexts,
				480, BigInteger.valueOf(3), BigInteger.valueOf(1));

		XWPFParagraph paragraph = document.createParagraph();
		paragraph.setNumID(numIDOne);
		XWPFRun run = paragraph.createRun();
		run.setText("第一章");

		paragraph = document.createParagraph();
		CTPPr ppr = ParagraphUtils.getParagraphCTPPr(paragraph);
		CTNumPr numPr = ParagraphUtils.getCTPPrCTNumPr(ppr);
		ParagraphUtils.setNumPr(numPr, numIDOne, BigInteger.valueOf(1));
		run = paragraph.createRun();
		run.setText("第一章第一節");
		
		paragraph = document.createParagraph();
		paragraph.setNumID(numIDOne);
		run = paragraph.createRun();
		run.setText("第二章");
		
		paragraph = document.createParagraph();
		ppr = ParagraphUtils.getParagraphCTPPr(paragraph);
		numPr = ParagraphUtils.getCTPPrCTNumPr(ppr);
		ParagraphUtils.setNumPr(numPr, numIDOne, BigInteger.valueOf(1));
		run = paragraph.createRun();
		run.setText("第二章第一節");

		document.write(out);
		out.close();
		System.out.println("applyingborder.docx written successully" + System.currentTimeMillis());

	}
}
