package Utils;

import java.math.BigInteger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEm;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHighlight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSignedTwipsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTextScale;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STEm;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;

public class HyperLinkUtils {
	/**
	 * @Description: 默認超連結格式
	 */
	public static void addParagraphTextHyperlinkBasicStyle(XWPFParagraph paragraph, String url, String text, String fontFamily,
			String fontSize, String colorVal, boolean isBlod, boolean isItalic, boolean isStrike) {
		addParagraphTextHyperlink(paragraph, url, text, fontFamily, fontSize, colorVal, isBlod, true, "0000FF",
				STUnderline.SINGLE, isItalic, isStrike, false, false, false, false, false, false, false, null, false,
				null, false, null, null, null, 0, 0, 0);
	}

	/**
	 * @Description: 設置超連結
	 * 
	 * @param paragraph		段落
	 * @param url			超連結
	 * @param text			欲綁定的文字
	 * @param fontFamily	字體
	 * @param fontSize		字體大小
	 * @param colorVal		顏色
	 * @param isBlod		粗體
	 * @param isUnderLine	底線
	 * @param underLineColor底線顏色
	 * @param underStyle	
	 * @param isItalic 		傾斜
	 * @param isStrike 		刪除線
	 * @param isDStrike
	 * @param isShadow 		陰影
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
	public static void addParagraphTextHyperlink(XWPFParagraph paragraph, String url, String text, String fontFamily,
			String fontSize, String colorVal, boolean isBlod, boolean isUnderLine, String underLineColor,
			STUnderline.Enum underStyle, boolean isItalic, boolean isStrike, boolean isDStrike, boolean isShadow,
			boolean isVanish, boolean isEmboss, boolean isImprint, boolean isOutline, boolean isEm, STEm.Enum emType,
			boolean isHightLight, STHighlightColor.Enum hightStyle, boolean isShd, STShd.Enum shdStyle, String shdColor,
			STVerticalAlignRun.Enum verticalAlign, int position, int spacingValue, int indent) {
		// Add the link as External relationship
		String id = paragraph.getDocument().getPackagePart()
				.addExternalRelationship(url, XWPFRelation.HYPERLINK.getRelation()).getId();
		// Append the link and bind it to the relationship
		CTHyperlink cLink = paragraph.getCTP().addNewHyperlink();
		cLink.setId(id);

		// Create the linked text
		CTText ctText = CTText.Factory.newInstance();
		ctText.setStringValue(text);
		CTR ctr = CTR.Factory.newInstance();
		CTRPr rpr = ctr.addNewRPr();

		if (StringUtils.isNotBlank(fontFamily)) {
			// 设置字体
			CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
			fonts.setAscii(fontFamily);
			fonts.setEastAsia(fontFamily);
			fonts.setHAnsi(fontFamily);
		}
		if (StringUtils.isNotBlank(fontSize)) {
			// 设置字体大小
			CTHpsMeasure sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(new BigInteger(fontSize));

			CTHpsMeasure szCs = rpr.isSetSzCs() ? rpr.getSzCs() : rpr.addNewSzCs();
			szCs.setVal(new BigInteger(fontSize));
		}
		// 设置超链接样式
		// 字体颜色
		if (StringUtils.isNotBlank(colorVal)) {
			CTColor color = CTColor.Factory.newInstance();
			color.setVal(colorVal);
			rpr.setColor(color);
		}
		// 加粗
		if (isBlod) {
			CTOnOff bCtOnOff = rpr.addNewB();
			bCtOnOff.setVal(STOnOff.TRUE);
		}
		// 下划线
		if (isUnderLine) {
			CTUnderline udLine = rpr.addNewU();
			udLine.setVal(underStyle);
			if(StringUtils.isNotBlank(underLineColor)) {
				udLine.setColor(underLineColor);
			}
		}
		// 倾斜
		if (isItalic) {
			CTOnOff iCtOnOff = rpr.addNewI();
			iCtOnOff.setVal(STOnOff.TRUE);
		}
		// 删除线
		if (isStrike) {
			CTOnOff sCtOnOff = rpr.addNewStrike();
			sCtOnOff.setVal(STOnOff.TRUE);
		}
		// 双删除线
		if (isDStrike) {
			CTOnOff dsCtOnOff = rpr.addNewDstrike();
			dsCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阴影
		if (isShadow) {
			CTOnOff shadowCtOnOff = rpr.addNewShadow();
			shadowCtOnOff.setVal(STOnOff.TRUE);
		}
		// 隐藏
		if (isVanish) {
			CTOnOff vanishCtOnOff = rpr.addNewVanish();
			vanishCtOnOff.setVal(STOnOff.TRUE);
		}
		// 浮凸
		if (isEmboss) {
			CTOnOff embossCtOnOff = rpr.addNewEmboss();
			embossCtOnOff.setVal(STOnOff.TRUE);
		}
		// 阴文
		if (isImprint) {
			CTOnOff isImprintCtOnOff = rpr.addNewImprint();
			isImprintCtOnOff.setVal(STOnOff.TRUE);
		}
		// 空心
		if (isOutline) {
			CTOnOff isOutlineCtOnOff = rpr.addNewOutline();
			isOutlineCtOnOff.setVal(STOnOff.TRUE);
		}
		// 着重号
		if (isEm) {
			CTEm em = rpr.addNewEm();
			em.setVal(emType);
		}
		// 突出显示文本
		if (isHightLight) {
			if (hightStyle != null) {
				CTHighlight hightLight = rpr.addNewHighlight();
				hightLight.setVal(hightStyle);
			}
		}
		if (isShd) {
			// 设置底纹
			CTShd shd = rpr.addNewShd();
			if (shdStyle != null) {
				shd.setVal(shdStyle);
			}
			if (shdColor != null) {
				shd.setColor(shdColor);
			}
		}
		// 上标下标
		if (verticalAlign != null) {
			rpr.addNewVertAlign().setVal(verticalAlign);
		}
		// 设置文本位置
		rpr.addNewPosition().setVal(new BigInteger(String.valueOf(position)));
		if (spacingValue != 0) {
			// 设置字符间距信息
			CTSignedTwipsMeasure ctSTwipsMeasure = rpr.addNewSpacing();
			ctSTwipsMeasure.setVal(new BigInteger(String.valueOf(spacingValue)));
		}
		// 设置字符间距缩进
		if (indent > 0) {
			CTTextScale paramCTTextScale = rpr.addNewW();
			paramCTTextScale.setVal(indent);
		}
		ctr.setTArray(new CTText[] { ctText });
		cLink.setRArray(new CTR[] { ctr });
	}
}
