package WordTest.WordTest;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

public class CreateWordSimplestNumberingList {
	public static void main(String[] args) throws Exception {
		
		XWPFDocument document = new XWPFDocument();

		XWPFStyles styles = document.createStyles();

		CTFonts fonts = CTFonts.Factory.newInstance();
		fonts.setAscii("times new roman");
		fonts.setHAnsi("times new roman");
		fonts.setEastAsia("times new roman");
		styles.setDefaultFonts(fonts);
		
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText("測試123:");
		run.setFontFamily("DFKai-SB");

		ArrayList<String> documentList = new ArrayList<String>(Arrays.asList(new String[] { "One", "Two", "Three" }));

		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		// Next we set the AbstractNumId. This requires care.
		// Since we are in a new document we can start numbering from 0.
		// But if we have an existing document, we must determine the next free number
		// first.
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));


		/// * Decimal list
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.CHINESE_COUNTING);
		cTLvl.addNewLvlText().setVal("%1、");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
		// */

		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);

		XWPFNumbering numbering = document.createNumbering();

		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);

		BigInteger numID = numbering.addNum(abstractNumID);

		for (String string : documentList) {
			paragraph = document.createParagraph();
			paragraph.setNumID(numID);
			run = paragraph.createRun();
			run.setText(string);
		}

		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setText("測試");
		run.addTab();
		run.setText("測試");
		document.write(new FileOutputStream("CreateWordSimplestNumberingList.docx"));
		document.close();
		System.out.println("createparagraph.docx written successfully");
	}
}
