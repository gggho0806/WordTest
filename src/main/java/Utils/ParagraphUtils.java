package Utils;

import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEm;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHighlight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSignedTwipsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTextScale;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STEm;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class ParagraphUtils {
	public static CTPPr getParagraphCTPPr(XWPFParagraph p) {
		CTPPr pPPr = null;
		if (p.getCTP() != null) {
			if (p.getCTP().getPPr() != null) {
				pPPr = p.getCTP().getPPr();
			} else {
				pPPr = p.getCTP().addNewPPr();
			}
		}
		return pPPr;
	}

	/**
	 * @param numPr
	 * @param numId
	 *            摘要編號
	 * @param ilvl
	 *            層級編號
	 */
	public static void setNumPr(CTNumPr numPr, BigInteger numId, BigInteger ilvl) {
		CTDecimalNumber number = CTDecimalNumber.Factory.newInstance();
		number.setVal(ilvl);
		numPr.setIlvl(number);
		number = CTDecimalNumber.Factory.newInstance();
		number.setVal(numId);
		numPr.setNumId(number);
	}

	public static CTNumPr getCTPPrCTNumPr(CTPPr ppr) {
		CTNumPr numPr = ppr.isSetNumPr() ? ppr.getNumPr() : ppr.addNewNumPr();
		return numPr;
	}

	public static CTRPr getRunCTRPr(XWPFParagraph p, XWPFRun pRun) {
		CTRPr pRpr = null;
		if (pRun.getCTR() != null) {
			pRpr = pRun.getCTR().getRPr();
			if (pRpr == null) {
				pRpr = pRun.getCTR().addNewRPr();
			}
		} else {
			pRpr = p.getCTP().addNewR().addNewRPr();
		}
		return pRpr;
	}

	public static XWPFRun getOrAddParagraphFirstRun(XWPFParagraph p, boolean isInsert, boolean isNewLine) {
		XWPFRun pRun = null;
		if (isInsert) {
			pRun = p.createRun();
		} else {
			if (p.getRuns() != null && p.getRuns().size() > 0) {
				pRun = p.getRuns().get(0);
			} else {
				pRun = p.createRun();
			}
		}
		if (isNewLine) {
			pRun.addBreak();
		}
		return pRun;
	}

	/**
	 * @param p
	 * @param pRun
	 * @param colorVal
	 * @param isBlod
	 * @param isUnderLine
	 * @param underLineColor
	 * @param underStyle
	 * @param isItalic
	 * @param isStrike
	 * @param isDStrike
	 * @param isShadow
	 * @param isVanish
	 * @param isEmboss
	 * @param isImprint
	 * @param isOutline
	 * @param isEm
	 * @param emType
	 * @param isHightLight
	 * @param hightStyle
	 * @param isShd
	 * @param shdStyle
	 * @param shdColor
	 * @param verticalAlign
	 * @param position
	 * @param spacingValue
	 * @param indent
	 */
	public static void setParagraphTextStyleInfo(XWPFParagraph p, XWPFRun pRun, String colorVal, boolean isBlod,
			boolean isUnderLine, String underLineColor, STUnderline.Enum underStyle, boolean isItalic, boolean isStrike,
			boolean isDStrike, boolean isShadow, boolean isVanish, boolean isEmboss, boolean isImprint,
			boolean isOutline, boolean isEm, STEm.Enum emType, boolean isHightLight, STHighlightColor.Enum hightStyle,
			boolean isShd, STShd.Enum shdStyle, String shdColor, VerticalAlign verticalAlign, int position,
			int spacingValue, int indent, int fontSize) {
		if (pRun == null) {
			return;
		}
		CTRPr pRpr = getRunCTRPr(p, pRun);
		if (colorVal != null) {
			pRun.setColor(colorVal);
		}
		// 设置字体样式
		// 加粗
		if (isBlod) {
			pRun.setBold(isBlod);
		}
		// 倾斜
		if (isItalic) {
			pRun.setItalic(isItalic);
		}
		// 删除线
		if (isStrike) {
			pRun.setStrikeThrough(isStrike);
		}
		// 双删除线
		if (isDStrike) {
			CTOnOff dsCtOnOff = pRpr.isSetDstrike() ? pRpr.getDstrike() : pRpr.addNewDstrike();
			dsCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阴影
		if (isShadow) {
			CTOnOff shadowCtOnOff = pRpr.isSetShadow() ? pRpr.getShadow() : pRpr.addNewShadow();
			shadowCtOnOff.setVal(STOnOff.TRUE);
		}
		// 隐藏
		if (isVanish) {
			CTOnOff vanishCtOnOff = pRpr.isSetVanish() ? pRpr.getVanish() : pRpr.addNewVanish();
			vanishCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阳文
		if (isEmboss) {
			CTOnOff embossCtOnOff = pRpr.isSetEmboss() ? pRpr.getEmboss() : pRpr.addNewEmboss();
			embossCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阴文
		if (isImprint) {
			CTOnOff isImprintCtOnOff = pRpr.isSetImprint() ? pRpr.getImprint() : pRpr.addNewImprint();
			isImprintCtOnOff.setVal(STOnOff.TRUE);
		}
		// 空心
		if (isOutline) {
			CTOnOff isOutlineCtOnOff = pRpr.isSetOutline() ? pRpr.getOutline() : pRpr.addNewOutline();
			isOutlineCtOnOff.setVal(STOnOff.TRUE);
		}
		// 着重号
		if (isEm) {
			CTEm em = pRpr.isSetEm() ? pRpr.getEm() : pRpr.addNewEm();
			em.setVal(emType);
		}
		// 设置下划线样式
		if (isUnderLine) {
			CTUnderline u = pRpr.isSetU() ? pRpr.getU() : pRpr.addNewU();
			if (underStyle != null) {
				u.setVal(underStyle);
			}
			if (underLineColor != null) {
				u.setColor(underLineColor);
			}
		}
		// 设置突出显示文本
		if (isHightLight) {
			if (hightStyle != null) {
				CTHighlight hightLight = pRpr.isSetHighlight() ? pRpr.getHighlight() : pRpr.addNewHighlight();
				hightLight.setVal(hightStyle);
			}
		}
		if (isShd) {
			// 设置底纹
			CTShd shd = pRpr.isSetShd() ? pRpr.getShd() : pRpr.addNewShd();
			if (shdStyle != null) {
				shd.setVal(shdStyle);
			}
			if (shdColor != null) {
				shd.setColor(shdColor);
			}
		}
		// 上标下标
		if (verticalAlign != null) {
			pRun.setSubscript(verticalAlign);
		}
		// 设置文本位置
		pRun.setTextPosition(position);
		if (spacingValue > 0) {
			// 设置字符间距信息
			CTSignedTwipsMeasure ctSTwipsMeasure = pRpr.isSetSpacing() ? pRpr.getSpacing() : pRpr.addNewSpacing();
			ctSTwipsMeasure.setVal(new BigInteger(String.valueOf(spacingValue)));
		}
		if (indent > 0) {
			CTTextScale paramCTTextScale = pRpr.isSetW() ? pRpr.getW() : pRpr.addNewW();
			paramCTTextScale.setVal(indent);
		}

		if (fontSize > 0) {
			pRun.setFontSize(fontSize);
		}
	}

	/**
	 * 垂直置中，必須在當頁字串都輸出後才可執行，因為當Paragraph加入SectPr後會自動換頁!! 換頁後垂直設定並不會影響下一個段落
	 * 
	 * @param paragraph
	 * @param vJcEnum
	 */
	public static void changeVerticalAlign(XWPFParagraph paragraph, STVerticalJc.Enum vJcEnum) {
		CTP ctp = paragraph.getCTP();
		CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
		CTSectPr sectPr = ppr.isSetSectPr() ? ppr.getSectPr() : ppr.addNewSectPr();
		CTVerticalJc vjc = CTVerticalJc.Factory.newInstance();
		vjc.setVal(vJcEnum);
		sectPr.setVAlign(vjc);
	}
}
