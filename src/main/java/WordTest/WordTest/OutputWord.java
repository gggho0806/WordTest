package WordTest.WordTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

public class OutputWord {
	public static void main(String[] args) throws IOException {
		// Blank Document
		XWPFDocument document = new XWPFDocument();

		// Write the Document in file system
		FileOutputStream out = new FileOutputStream(new File("output.docx"));

		createDefaultHeader(document, "testtesttesttesttest");
		createDefaultFooter(document, "test");
		// 設定統一字體
		XWPFStyles styles = document.createStyles();
		CTFonts fonts = CTFonts.Factory.newInstance();
		fonts.setAscii("Times new roman");
		fonts.setHAnsi("標楷體");
		fonts.setEastAsia("標楷體");
		styles.setDefaultFonts(fonts);
		XWPFParagraph paragraph = document.createParagraph();

		addNomalParagraph(paragraph, "標題1", ParagraphAlignment.CENTER, TextAlignment.CENTER, 20, false, false, null,
				null, BreakType.TEXT_WRAPPING);

		addNomalParagraph(paragraph, "標題2", ParagraphAlignment.CENTER, TextAlignment.BOTTOM, 20, false, false, null,
				null, BreakType.TEXT_WRAPPING);

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
	public static void addNomalParagraph(XWPFParagraph paragraph, String content, ParagraphAlignment hAlignment,
			TextAlignment vAlignment, Integer fontSize, Boolean isItalic, Boolean isBold,
			UnderlinePatterns underLinePattern, String rgbString, BreakType breakType) {

		paragraph.setAlignment(hAlignment);
		paragraph.setVerticalAlignment(vAlignment);

		XWPFRun run = paragraph.createRun();
		if (fontSize != null)
			run.setFontSize(fontSize);
		if (isItalic != null)
			run.setItalic(isItalic);
		if (isBold != null)
			run.setBold(isBold);
		if (underLinePattern != null)
			run.setUnderline(underLinePattern);
		if (rgbString != null)
			run.setColor(rgbString);

		run.setText(content);
		run.addBreak(breakType);
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
	 * 建立預設頁首
	 *
	 * @param docx XWPFDocument文件物件
	 * @param text 頁首文字
	 * @return 返回文件幫助類物件，可用於方法鏈呼叫
	 * @throws XmlException XML異常
	 * @throws IOException IO異常
	 * @throws InvalidFormatException 非法格式異常
	 * @throws FileNotFoundException 找不到文件異常
	 */
	public static void createDefaultHeader(final XWPFDocument docx, final String text) throws IOException{
		CTP ctp = CTP.Factory.newInstance();
	    XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
	    paragraph.setAlignment(ParagraphAlignment.RIGHT);
	    ctp.addNewR().addNewT().setStringValue(text);
	    ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.DEFAULT);
	    CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
	    XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
	    XWPFHeader header = policy.createHeader(STHdrFtr.FIRST, new XWPFParagraph[] { paragraph });
	    header.setXWPFDocument(docx);
	}
	
	/**
	 * @param docx XWPFDocument文件物件
	 * @param text 頁尾內容
	 * @throws IOException IO異常
	 */
	public static void createDefaultFooter(final XWPFDocument docx, final String text) throws IOException {
		CTP ctp = CTP.Factory.newInstance();
	    XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
	    ctp.addNewR().addNewT().setStringValue(text);
	    ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
	    CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
	    XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
	    XWPFFooter footer = policy.createFooter(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
	    footer.setXWPFDocument(docx);
	}
}
