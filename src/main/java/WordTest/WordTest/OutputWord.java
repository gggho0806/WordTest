package WordTest.WordTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;

public class OutputWord {
	public static void main(String[] args) throws IOException {
		// Blank Document
		XWPFDocument document = new XWPFDocument();
		
		// Write the Document in file system
		FileOutputStream out = new FileOutputStream(new File("output.docx"));
		
		// 設定統一字體
		XWPFStyles styles = document.createStyles();
		CTFonts fonts = CTFonts.Factory.newInstance();
		fonts.setAscii("Times new roman");
		fonts.setHAnsi("標楷體");
		fonts.setEastAsia("標楷體");
		styles.setDefaultFonts(fonts);
		
		XWPFParagraph paragraph = document.createParagraph();
		
		addNomalParagraph(paragraph, "台灣積體電路製造股份有限公司及子公司", ParagraphAlignment.CENTER
				, TextAlignment.CENTER, 20, false, false, null, null, BreakType.TEXT_WRAPPING);
		
		addNomalParagraph(paragraph, "合併財務報告暨會計師核閱報告", ParagraphAlignment.MEDIUM_KASHIDA
				, TextAlignment.CENTER, 20, false, false, null, null, BreakType.TEXT_WRAPPING);
		
		addNomalParagraph(paragraph, "民國107 及106 年第2 季", ParagraphAlignment.CENTER
				, TextAlignment.BOTTOM, 15, false, false, null, null, BreakType.TEXT_WRAPPING);

		
		
		document.write(out);
		out.close();
		System.out.println("applyingborder.docx written successully"+System.currentTimeMillis());
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
	public static void addNomalParagraph(XWPFParagraph paragraph, String content, ParagraphAlignment hAlignment, TextAlignment vAlignment, Integer fontSize,
			Boolean isItalic, Boolean isBold, UnderlinePatterns underLinePattern, String rgbString, BreakType breakType) {
		
		paragraph.setAlignment(hAlignment);
		paragraph.setVerticalAlignment(vAlignment);
		
		XWPFRun run=paragraph.createRun();
		if(fontSize != null)
			run.setFontSize(fontSize);
		if(isItalic != null)
			run.setItalic(isItalic);
		if(isBold != null)
			run.setBold(isBold);
		if(underLinePattern != null)
			run.setUnderline(underLinePattern);
		if(rgbString != null)
			run.setColor(rgbString);
		
		run.setText(content);
		run.addBreak(breakType);
	}
}
