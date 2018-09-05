package Style;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;

public class TextStyle {
	public TextStyle() {
		
	}
	public TextStyle(String content, ParagraphAlignment hAlignment, TextAlignment vAlignment, Integer fontSize,
			Boolean isItalic, Boolean isBold, UnderlinePatterns underLinePattern, String rgbString,
			BreakType breakType) {
		super();
		this.content = content;
		this.hAlignment = hAlignment;
		this.vAlignment = vAlignment;
		this.fontSize = fontSize;
		this.isItalic = isItalic;
		this.isBold = isBold;
		this.underLinePattern = underLinePattern;
		this.rgbString = rgbString;
		this.breakType = breakType;
	}
	String content;
	ParagraphAlignment hAlignment;
	TextAlignment vAlignment;
	Integer fontSize;
	Boolean isItalic;
	Boolean isBold;
	UnderlinePatterns underLinePattern;
	String rgbString;
	BreakType breakType;
	public String getContent() {
		return content;
	}
	public void setContent(String content) {
		this.content = content;
	}
	public ParagraphAlignment gethAlignment() {
		return hAlignment;
	}
	public void sethAlignment(ParagraphAlignment hAlignment) {
		this.hAlignment = hAlignment;
	}
	public TextAlignment getvAlignment() {
		return vAlignment;
	}
	public void setvAlignment(TextAlignment vAlignment) {
		this.vAlignment = vAlignment;
	}
	public Integer getFontSize() {
		return fontSize;
	}
	public void setFontSize(Integer fontSize) {
		this.fontSize = fontSize;
	}
	public Boolean getIsItalic() {
		return isItalic;
	}
	public void setIsItalic(Boolean isItalic) {
		this.isItalic = isItalic;
	}
	public Boolean getIsBold() {
		return isBold;
	}
	public void setIsBold(Boolean isBold) {
		this.isBold = isBold;
	}
	public UnderlinePatterns getUnderLinePattern() {
		return underLinePattern;
	}
	public void setUnderLinePattern(UnderlinePatterns underLinePattern) {
		this.underLinePattern = underLinePattern;
	}
	public String getRgbString() {
		return rgbString;
	}
	public void setRgbString(String rgbString) {
		this.rgbString = rgbString;
	}
	public BreakType getBreakType() {
		return breakType;
	}
	public void setBreakType(BreakType breakType) {
		this.breakType = breakType;
	}
}
