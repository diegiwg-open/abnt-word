/**
 * Quick Selection Formatter
 * Applies instant ABNT styles to the currently selected text or paragraph.
 */

import { ABNT_CONSTANTS } from '../utils/abntRules.js';
import { wordBridge } from '../wordApi.js';

export class SelectionFormatter {
  /**
   * Formats the selection as standard ABNT body text
   * - 12pt font, 1.5 line spacing, 1.25 cm first-line indent, Justified, 0pt before/after
   */
  static async formatBodyText(fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: ABNT_CONSTANTS.FONTS.SIZES.BODY,
      bold: false,
      italic: false,
      alignment: ABNT_CONSTANTS.ALIGNMENT.JUSTIFIED,
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.BODY, // 1.5
      firstLineIndent: ABNT_CONSTANTS.INDENTATION.FIRST_LINE_PT, // 1.25 cm
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: ABNT_CONSTANTS.PARAGRAPH_SPACING.BODY_BEFORE_PT,
      spaceAfter: ABNT_CONSTANTS.PARAGRAPH_SPACING.BODY_AFTER_PT,
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }

  /**
   * Formats the selection as a Long Direct Quotation (> 3 lines) – NBR 10520
   * - 4.0 cm left indent from the left margin
   * - 10pt font size
   * - Single line spacing
   * - No first-line indent
   * - Justified alignment
   * - 6pt space before and after
   */
  static async formatLongCitation(fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: ABNT_CONSTANTS.FONTS.SIZES.LONG_CITATION, // 10 pt
      bold: false,
      italic: false,
      alignment: ABNT_CONSTANTS.ALIGNMENT.JUSTIFIED,
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.SINGLE, // 1.0 (single)
      firstLineIndent: 0,
      leftIndent: ABNT_CONSTANTS.INDENTATION.LONG_CITATION_LEFT_PT, // 4.0 cm (113.39 pt)
      rightIndent: 0,
      spaceBefore: ABNT_CONSTANTS.PARAGRAPH_SPACING.LONG_CITATION_BEFORE_PT, // 6 pt
      spaceAfter: ABNT_CONSTANTS.PARAGRAPH_SPACING.LONG_CITATION_AFTER_PT,   // 6 pt
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }

  /**
   * Formats the selection as a Figure / Table Caption
   * - 10pt, Centered, Single line spacing
   */
  static async formatCaption(fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: ABNT_CONSTANTS.FONTS.SIZES.CAPTION, // 10 pt
      bold: false,
      italic: false,
      alignment: ABNT_CONSTANTS.ALIGNMENT.CENTERED,
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.SINGLE,
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: 6,
      spaceAfter: 3,
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }

  /**
   * Formats the selection as a Footnote or Source attribution
   * - 10pt, Left-aligned, Single line spacing
   */
  static async formatFootnoteOrSource(fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: ABNT_CONSTANTS.FONTS.SIZES.FOOTNOTE, // 10 pt
      bold: false,
      italic: false,
      alignment: ABNT_CONSTANTS.ALIGNMENT.LEFT,
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.SINGLE,
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: 3,
      spaceAfter: 6,
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }

  /**
   * Formats the selection as a Bibliographic Reference item (NBR 6023)
   * - 12pt, Left-aligned, Single line spacing, 6pt after
   */
  static async formatReferenceItem(fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: ABNT_CONSTANTS.FONTS.SIZES.BODY,
      bold: false,
      italic: false,
      alignment: ABNT_CONSTANTS.ALIGNMENT.LEFT,
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.SINGLE,
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: 0,
      spaceAfter: ABNT_CONSTANTS.PARAGRAPH_SPACING.REFERENCE_AFTER_PT, // 6 pt
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }
}
