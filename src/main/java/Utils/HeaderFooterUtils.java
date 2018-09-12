package Utils;

import java.io.IOException;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class HeaderFooterUtils {
	
	/**
	 * 預設頁首，置中且為每一頁
	 * @param document
	 * @param text
	 * @throws IOException
	 */
	public static void createDefaultNormaltHeader(XWPFDocument document, String text) throws IOException {
		createNormaltHeader(document, text, ParagraphAlignment.CENTER, HeaderFooterType.DEFAULT);
	}
		
	/**
	 * @param document	XWPFDocument物件
	 * @param text		頁首文字
	 * @param aligment
	 * @param hfType
	 * @throws IOException
	 */
	public static void createNormaltHeader(XWPFDocument document, String text, ParagraphAlignment aligment, HeaderFooterType hfType) throws IOException {
		XWPFHeader header = document.createHeader(hfType);
		XWPFParagraph paragraph = header.createParagraph();
		paragraph.setAlignment(aligment);
		XWPFRun run = paragraph.createRun();
		run.setText(text);
	}
	
	/**
	 * 自動生成頁碼，格式為Page ? Of ?<br>
	 * 第一個問號為當前頁碼，第二個問號為總頁數
	 * @param document	文件主體
	 * @throws Exception
	 */
	public static void createPageNumFooter(XWPFDocument document,ParagraphAlignment alignment) throws Exception {
		// 建立頁尾
		XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
		XWPFParagraph paragraph = footer.createParagraph();
		paragraph.setAlignment(alignment);

		// 生成頁碼開頭
		XWPFRun run = paragraph.createRun();
		run.setText("Page ");
		paragraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
		run = paragraph.createRun();
		run.setText(" of ");
		paragraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");

	}
}
