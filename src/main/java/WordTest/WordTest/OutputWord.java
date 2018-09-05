package WordTest.WordTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

import Style.TableStyle;
import Style.TextStyle;

public class OutputWord {
	public static void main(String[] args) throws Exception {
		// Blank Document
		XWPFDocument document = new XWPFDocument();

		// Write the Document in file system
		FileOutputStream out = new FileOutputStream(new File("output.docx"));
		createPageNumFooter(document);
		createNormaltHeader(document, "testtesttesttesttest");

		// 設定統一字體
		XWPFStyles styles = document.createStyles();
		CTFonts fonts = CTFonts.Factory.newInstance();
		fonts.setAscii("Times new roman");
		fonts.setHAnsi("標楷體");
		fonts.setEastAsia("標楷體");
		styles.setDefaultFonts(fonts);
		XWPFParagraph paragraph = document.createParagraph();
		TextStyle style = new TextStyle("標題1", ParagraphAlignment.CENTER, TextAlignment.CENTER, 20, false, false, null,
				null, BreakType.TEXT_WRAPPING);
		addNomalParagraph(paragraph, style);
		style = new TextStyle("標題2", ParagraphAlignment.CENTER, TextAlignment.BOTTOM, 20, false, false, null, null,
				BreakType.PAGE);
		addNomalParagraph(paragraph, style);

		document.write(out);
		out.close();
		System.out.println("applyingborder.docx written successully" + System.currentTimeMillis());
	}

	/**
	 * @param paragraph
	 * @param content
	 * @param hAlignment
	 * @param vAlignment
	 * @param fontSize
	 * @param isItalic
	 * @param isBold
	 * @param underLinePattern
	 * @param rgbString
	 * @param breakType
	 */
	public static void addNomalParagraph(XWPFParagraph paragraph, TextStyle style) {

		paragraph.setAlignment(style.gethAlignment());
		paragraph.setVerticalAlignment(style.getvAlignment());

		XWPFRun run = paragraph.createRun();
		if (style.getFontSize() != null)
			run.setFontSize(style.getFontSize());
		if (style.getIsItalic() != null)
			run.setItalic(style.getIsItalic());
		if (style.getIsBold() != null)
			run.setBold(style.getIsBold());
		if (style.getUnderLinePattern() != null)
			run.setUnderline(style.getUnderLinePattern());
		if (style.getRgbString() != null)
			run.setColor(style.getRgbString());

		run.setText(style.getContent());
		run.addBreak(style.getBreakType());
	}

	public static void createTable(XWPFDocument document, TableStyle style) {
		// create table
		XWPFTable table = document.createTable(style.getRowNum(), style.getColumnNum());
		table.setWidth(style.getTableWidth());
		String[][] contentMatrix = style.getContent();
		for (int i = 0; i < style.getRowNum(); i++) {
			XWPFTableRow row = table.getRow(i);
			for (int j = 0; j < style.getColumnNum(); j++) {
				row.getCell(j).setText(contentMatrix[i][j]);
			}
		}
	}

	/**
	 * 設定頁邊距 (word中1釐米約等於567)
	 * 
	 * @param document
	 * @param left
	 * @param top
	 * @param right
	 * @param bottom
	 */
	public static void setDocumentMargin(XWPFDocument document, String left, String top, String right, String bottom) {
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		CTPageMar ctpagemar = sectPr.addNewPgMar();
		if (StringUtils.isNotBlank(left)) {
			ctpagemar.setLeft(new BigInteger(left));
		}
		if (StringUtils.isNotBlank(top)) {
			ctpagemar.setTop(new BigInteger(top));
		}
		if (StringUtils.isNotBlank(right)) {
			ctpagemar.setRight(new BigInteger(right));
		}
		if (StringUtils.isNotBlank(bottom)) {
			ctpagemar.setBottom(new BigInteger(bottom));
		}
	}

	/**
	 * @param document
	 *            XWPFDocument文件物件
	 * @param text
	 *            頁首文字
	 * @throws IOException
	 */
	public static void createNormaltHeader(final XWPFDocument document, final String text) throws IOException {
		XWPFHeader header = document.createHeader(HeaderFooterType.FIRST);
		XWPFParagraph paragraph = header.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.RIGHT);
		XWPFRun run = paragraph.createRun();
		run.setText(text);
	}

	/**
	 * @param docx
	 *            XWPFDocument文件物件
	 * @param text
	 *            頁尾內容
	 * @throws IOException
	 *             IO異常
	 */
	public static void createDefaultFooter(final XWPFDocument docx, final String text) throws IOException {
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
		ctp.addNewR().addNewT().setStringValue(text);
		ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
		CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr()
				: docx.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
		XWPFFooter footer = policy.createFooter(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
		footer.setXWPFDocument(docx);
	}

	public static void createPageNumFooter(XWPFDocument document) throws Exception {

		// 如果頁頭有設定過FIRST，會造成頁尾設定DEFAULT時，首頁會吃不到，所以需額外設定首頁
		// 建立頁尾
		XWPFFooter footer = document.createFooter(HeaderFooterType.FIRST);
		XWPFParagraph paragraph = footer.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		paragraph.setVerticalAlignment(TextAlignment.CENTER);

		// 生成頁碼開頭
		XWPFRun run = paragraph.createRun();
		run.setText("Page ");
		paragraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
		run = paragraph.createRun();
		run.setText(" of ");
		paragraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");

		// 建立頁尾
		footer = document.createFooter(HeaderFooterType.DEFAULT);
		paragraph = footer.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		paragraph.setVerticalAlignment(TextAlignment.CENTER);

		// 生成頁碼開頭
		run = paragraph.createRun();
		run.setText("-");
		paragraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
		run = paragraph.createRun();
		run.setText("-");
		paragraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");

	}
}
