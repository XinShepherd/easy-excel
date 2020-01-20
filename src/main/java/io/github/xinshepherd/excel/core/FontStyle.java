package io.github.xinshepherd.excel.core;

import org.apache.poi.ss.usermodel.Font;
/**
 * 自定义字体
 *
 * @see org.apache.poi.ss.usermodel.Font
 * @author Fuxin
 * @since 1.1.0
 */
public interface FontStyle {

    /**
     * get the name for the font (i.e. Arial)
     * default auto use System font name.
     * @return String representing the name of the font to use
     */
    default String getFontName() {
        return null;
    }

    /**
     * get whether to use italics or not
     * @return italics or not
     */
    default boolean getItalic() {
        return Boolean.FALSE;
    }

    /**
     * get the color for the font
     * @return color to use
     * @see org.apache.poi.hssf.usermodel.HSSFPalette#getColor(short)
     */
    default short getColor() {
        return -1;
    }

    /**
     * get whether to use bold or not
     *
     * @see Font#getBold()
     * @return bold or not
     */
    default boolean getBold() {
        return Boolean.FALSE;
    }

    /**
     * Get the font height in unit's of 1/20th of a point.
     * <p>
     * For many users, the related {@link #getFontHeightInPoints()}
     *  will be more helpful, as that returns font heights in the
     *  more familiar points units, eg 10, 12, 14.

     * @return short - height in 1/20ths of a point
     * @see #getFontHeightInPoints()
     * @see Font#getFontHeight()
     */
    default short getFontHeight() {
        return -1;
    }

    /**
     * Get the font height in points.
     * <p>
     * This will return the same font height that is shown in Excel,
     *  such as 10 or 14 or 28.
     * @return short - height in the familiar unit of measure - points
     * @see #getFontHeight()
     * @see Font#getFontHeightInPoints()
     */
    default short getFontHeightInPoints(){
        return -1;
    }

    /**
     * get type of text underlining to use
     * @return underlining type
     * @see Font#getUnderline()
     */
    default byte getUnderline() {
        return Font.U_NONE;
    }

    /**
     * get normal,super or subscript.
     * @return offset type to use (none,super,sub)
     * @see Font#getTypeOffset()
     */
    default short getTypeOffset() {
        return Font.SS_NONE;
    }

    /**
     * get character-set to use.
     * @return character-set
     * @see Font#getCharSet()
     */
    default int getCharSet(){
        return Font.DEFAULT_CHARSET;
    }

    /**
     * get whether to use a strikeout horizontal line through the text or not
     * @see Font#getStrikeout()
     * @return strikeout or not
     */
    default boolean getStrikeout() {
        return false;
    }
}
