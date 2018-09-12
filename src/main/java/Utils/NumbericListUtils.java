package Utils;

import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

public class NumbericListUtils {
	
	
	/**
	 * @param document
	 * @param stNumberFormat
	 * @param lvlTexts
	 * @param baseMarginLeft	每一層編號都往右移多少
	 * @param abstractNumId		摘要的編號
	 * @param start
	 * @return
	 */
	public static BigInteger addAbstractNum(XWPFDocument document, STNumberFormat.Enum stNumberFormat, String[] lvlTexts,
			int baseMarginLeft, BigInteger abstractNumId, BigInteger start) {
		//
		CTAbstractNum ctAbstractNum = CTAbstractNum.Factory.newInstance();
		/*
		 * 分成多個AbstractNum，要注意的是如果該編號可對應到存在的AbstractNum，
		 * 那之後的變動會對此AbstractNum造成影響，但不包括重置編號
		 */
		ctAbstractNum.setAbstractNumId(abstractNumId);
		
		for(int i=0;i<lvlTexts.length;i++) {
			// 建立每一層的樣式
			String replace = "%"+(i+1);
			String lvlPattern = lvlTexts[i].replace("{num}", replace);
			// 取得層級
			CTLvl ctLvl = getCTLvl(ctAbstractNum, i);
			// 設定層級的編號
			BigInteger lvlNum = BigInteger.valueOf(i);
			ctLvl.setIlvl(lvlNum);
			// 設定編號數字部分的樣式
			ctLvl.addNewNumFmt().setVal(stNumberFormat);
			// 設定整體編號樣式，像是%1、為一、二、...
			ctLvl.addNewLvlText().setVal(lvlPattern);
			// 設定起始編號
			ctLvl.addNewStart().setVal(start);
			// 設定編號與左邊的邊距
			CTPPr ppr = getCTLvlCTPPr(ctLvl);
			CTInd ind = getCTPPrCTInd(ppr);
			BigInteger marginLeft = BigInteger.valueOf(i * baseMarginLeft);
			ind.setLeft(marginLeft);
		}
		

		// 將編號設定加入word
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(ctAbstractNum);
		XWPFNumbering numbering = getXWPFNumbering(document);

		// 取得編號設定的起始號碼，每當paragraph加入此號碼，會自動往上加1
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);

		return numID;
	}
	
	public static CTInd getCTPPrCTInd(CTPPr ppr) {
		CTInd ind = ppr.isSetInd() ? ppr.getInd() : ppr.addNewInd();
		return ind;
	}
	
	public static CTPPr getCTLvlCTPPr(CTLvl ctLvl) {
		CTPPr ppr = ctLvl.getPPr() == null ? ctLvl.addNewPPr() : ctLvl.getPPr();
		return ppr;
	}
	
	public static CTLvl getCTLvl(CTAbstractNum ctAbstractNum, int index) {
		CTLvl ctLvl = ctAbstractNum.getLvlList().size()<=index?ctAbstractNum.addNewLvl() : ctAbstractNum.getLvlArray(index);
		return ctLvl;
	}
	
	public static XWPFNumbering getXWPFNumbering(XWPFDocument document) {
		XWPFNumbering numbering = null;
		if (document.getNumbering() == null) {
			numbering = document.createNumbering();
		} else {
			numbering = document.getNumbering();
		}
		return numbering;
	}
}
