/**
 * Official ABNT Constants and Rules for Academic Papers
 * Based on standards:
 * - NBR 14724 (Academic work - Presentation)
 * - NBR 6023 (References - Elaboration)
 * - NBR 10520:2023 (Citations in documents)
 * - NBR 6024 (Progressive numbering of sections)
 * - NBR 6027 (Table of Contents / Summary)
 * - NBR 6028 (Abstract)
 */

export const ABNT_CONSTANTS = {
  // Unit Conversion (1 inch = 2.54 cm = 72 pt)
  CM_TO_POINTS: 72 / 2.54, // ~28.3465 pt per cm

  // Official Margins (NBR 14724)
  MARGINS: {
    TOP_CM: 3.0,
    LEFT_CM: 3.0,
    BOTTOM_CM: 2.0,
    RIGHT_CM: 2.0,
    TOP_PT: 3.0 * (72 / 2.54), // 85.04 pt
    LEFT_PT: 3.0 * (72 / 2.54), // 85.04 pt
    BOTTOM_PT: 2.0 * (72 / 2.54), // 56.69 pt
    RIGHT_PT: 2.0 * (72 / 2.54), // 56.69 pt
  },

  // Allowed Fonts
  FONTS: {
    ARIAL: 'Arial',
    TIMES: 'Times New Roman',
    SIZES: {
      BODY: 12,
      TITLE: 12,
      LONG_CITATION: 10,
      FOOTNOTE: 10,
      CAPTION: 10,
      PAGE_NUMBER: 10,
      TABLE_DATA: 10,
    },
    COLOR_BLACK: '#000000',
  },

  // Line Spacing
  LINE_SPACING: {
    BODY: 1.5, // 1.5 lines for main body text
    SINGLE: 1.0, // Single spacing for blockquotes, footnotes, references, captions
    DOUBLE: 2.0,
  },

  // Paragraph Indentation
  INDENTATION: {
    FIRST_LINE_CM: 1.25, // 1.25 cm first line indent
    FIRST_LINE_PT: 1.25 * (72 / 2.54), // 35.43 pt
    LONG_CITATION_LEFT_CM: 4.0, // 4.0 cm left indent for block quotations
    LONG_CITATION_LEFT_PT: 4.0 * (72 / 2.54), // 113.39 pt
    FOLHA_ROSTO_LEFT_CM: 8.0, // 8.0 cm left indent for title page nature note
    FOLHA_ROSTO_LEFT_PT: 8.0 * (72 / 2.54), // 226.77 pt
    NONE: 0,
  },

  // Spacing Before and After Paragraphs (in points)
  PARAGRAPH_SPACING: {
    BODY_BEFORE_PT: 0,
    BODY_AFTER_PT: 0,
    LONG_CITATION_BEFORE_PT: 6,
    LONG_CITATION_AFTER_PT: 6,
    HEADING_BEFORE_PT: 12,
    HEADING_AFTER_PT: 12,
    REFERENCE_AFTER_PT: 6, // 1 single space between references
  },

  // Text Alignment (Word API values)
  ALIGNMENT: {
    JUSTIFIED: 'Justified',
    LEFT: 'Left',
    CENTERED: 'Centered',
    RIGHT: 'Right',
  },

  // Section Heading Levels (NBR 6024)
  HEADINGS: {
    PRIMARY: {
      level: 1,
      name: 'Seção Primária',
      example: '1 INTRODUÇÃO',
      bold: true,
      italic: false,
      uppercase: true,
      size: 12,
      pageBreakBefore: true,
    },
    SECONDARY: {
      level: 2,
      name: 'Seção Secundária',
      example: '1.1 OBJETIVOS',
      bold: false,
      italic: false,
      uppercase: true,
      size: 12,
      pageBreakBefore: false,
    },
    TERTIARY: {
      level: 3,
      name: 'Seção Terciária',
      example: '1.1.1 Objetivo Geral',
      bold: true,
      italic: false,
      uppercase: false,
      size: 12,
      pageBreakBefore: false,
    },
    QUATERNARY: {
      level: 4,
      name: 'Seção Quaternária',
      example: '1.1.1.1 Detalhe',
      bold: false,
      italic: false,
      uppercase: false,
      size: 12,
      pageBreakBefore: false,
    },
    QUINARY: {
      level: 5,
      name: 'Seção Quinária',
      example: '1.1.1.1.1 Subdetalhe',
      bold: false,
      italic: true,
      uppercase: false,
      size: 12,
      pageBreakBefore: false,
    },
  },
};

/**
 * Converts centimeters to points
 * @param {number} cm
 * @returns {number} points
 */
export function cmToPt(cm) {
  return cm * ABNT_CONSTANTS.CM_TO_POINTS;
}

/**
 * Converts points to centimeters
 * @param {number} pt
 * @returns {number} centimeters
 */
export function ptToCm(pt) {
  return pt / ABNT_CONSTANTS.CM_TO_POINTS;
}

/**
 * Rounds a number to a specified number of decimal places
 * @param {number} num
 * @param {number} decimals
 * @returns {number}
 */
export function roundTo(num, decimals = 2) {
  const factor = Math.pow(10, decimals);
  return Math.round(num * factor) / factor;
}
