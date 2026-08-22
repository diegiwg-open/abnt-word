/**
 * Section Headings Formatter (ABNT NBR 6024)
 * Standardizes progressive numbering and typography for heading levels:
 * - 1 PRIMARY SECTION (Bold, All Caps)
 * - 1.1 SECONDARY SECTION (No Bold, All Caps)
 * - 1.1.1 Tertiary Section (Bold, Title Case)
 * - 1.1.1.1 Quaternary Section (No Bold, Title Case)
 * - 1.1.1.1.1 Quinary Section (Italic, Title Case)
 */

import { ABNT_CONSTANTS } from '../utils/abntRules.js';
import { wordBridge } from '../wordApi.js';

export class HeadingsFormatter {
  /**
   * Applies the formatting for the specified heading level
   * @param {number} level - Heading level (1 to 5)
   * @param {string} fontName - Font name ('Arial' or 'Times New Roman')
   */
  static async formatHeadingLevel(level, fontName = ABNT_CONSTANTS.FONTS.ARIAL) {
    let headingConfig;

    switch (level) {
      case 1: headingConfig = ABNT_CONSTANTS.HEADINGS.PRIMARY; break;
      case 2: headingConfig = ABNT_CONSTANTS.HEADINGS.SECONDARY; break;
      case 3: headingConfig = ABNT_CONSTANTS.HEADINGS.TERTIARY; break;
      case 4: headingConfig = ABNT_CONSTANTS.HEADINGS.QUATERNARY; break;
      case 5: headingConfig = ABNT_CONSTANTS.HEADINGS.QUINARY; break;
      default: headingConfig = ABNT_CONSTANTS.HEADINGS.PRIMARY;
    }

    return await wordBridge.formatSelection({
      fontName: fontName,
      fontSize: headingConfig.size, // Always 12pt in ABNT
      bold: headingConfig.bold,
      italic: headingConfig.italic,
      alignment: ABNT_CONSTANTS.ALIGNMENT.LEFT,
      lineSpacing: ABNT_CONSTANTS.LINE_SPACING.BODY, // 1.5
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: ABNT_CONSTANTS.PARAGRAPH_SPACING.HEADING_BEFORE_PT, // 12 pt
      spaceAfter: ABNT_CONSTANTS.PARAGRAPH_SPACING.HEADING_AFTER_PT,   // 12 pt
      color: ABNT_CONSTANTS.FONTS.COLOR_BLACK,
    });
  }

  /**
   * Removes erroneous dots after section numbers in heading text
   * e.g. "1. INTRODUÇÃO" → "1 INTRODUÇÃO", "1.1. OBJETIVO" → "1.1 OBJETIVO"
   * @param {string} text - Raw heading text
   * @returns {string} Sanitized heading text
   */
  static sanitizeHeadingText(text) {
    return text.replace(/^(\d+(\.\d+)*)\.\ +/, '$1 ');
  }
}
