package WordTest.WordTest;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

public class ParagraphReader {
	public static void main(String[] args) {
		try {
			FileInputStream fis = new FileInputStream("test.docx");
			XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));

			List<XWPFParagraph> paragraphList = xdoc.getParagraphs();
			for (XWPFParagraph paragraph : paragraphList) {

				System.out.println("Text:"+paragraph.getText());
				System.out.println("HAlignment:"+paragraph.getAlignment());
				System.out.println("VAlignment:"+paragraph.getVerticalAlignment());
				System.out.println("Run Size:"+paragraph.getRuns().size());
				System.out.println("Style:"+paragraph.getStyle());
				System.out.println("BorderBetween:"+paragraph);
				System.out.println("********************************************************************");
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
}
