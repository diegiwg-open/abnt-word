/**
 * Bibliographic References Formatter (ABNT NBR 6023)
 * - Left-aligned
 * - Single line spacing
 * - Separated by one single line space (6pt paragraph after)
 * - Strict alphabetical ordering
 */

import { ABNT_CONSTANTS } from '../utils/abntRules.js';
import { wordBridge } from '../wordApi.js';

export class ReferencesFormatter {
  /**
   * Formats the selected paragraphs as ABNT references
   * @param {string} fontName - Font name ('Arial' or 'Times New Roman')
   */
  static async formatReferencesSelection(fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: ABNT_CONSTANTS.FONTS.SIZES.BODY, // 12 pt
      bold: false,
      italic: false,
      alignment: ABNT_CONSTANTS.ALIGNMENT.LEFT, // Left-aligned (NBR 6023)
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.SINGLE, // Single
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: 0,
      spaceAfter: ABNT_CONSTANTS.PARAGRAPH_SPACING.REFERENCE_AFTER_PT, // 6 pt
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }

  /**
   * Sorts the selected reference paragraphs alphabetically (A-Z)
   */
  static async sortSelectedReferences() {
    return await wordBridge.sortReferences();
  }
}
