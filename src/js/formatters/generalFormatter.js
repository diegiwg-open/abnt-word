/**
 * One-Click General Formatter (ABNT NBR 14724)
 * Applies all structural formatting to the entire document:
 * - ABNT margins (Top: 3cm, Left: 3cm, Bottom: 2cm, Right: 2cm)
 * - Typography (Arial or Times New Roman 12pt, black)
 * - 1.5 line spacing
 * - 1.25 cm first-line indent
 * - Justified text alignment
 */

import { ABNT_CONSTANTS } from '../utils/abntRules.js';
import { wordBridge } from '../wordApi.js';

export class GeneralFormatter {
  /**
   * Runs full ABNT formatting on the entire document
   * @param {Object} options - { fontName: 'Arial' | 'Times New Roman' }
   */
  static async formatFullDocument(options = {}) {
    const fontName = options.fontName || ABNT_CONSTANTS.FONTS.ARIAL;

    try {
      // 1. Apply ABNT margins
      const marginsResult = await wordBridge.applyAbntMargins();

      // 2. Apply typography and paragraph formatting to the whole document
      const docResult = await wordBridge.formatEntireDocument({
        fontName: fontName,
        fontSize: ABNT_CONSTANTS.FONTS.SIZES.BODY,
        lineSpacing: ABNT_CONSTANTS.LINE_SPACING.BODY, // 1.5
        firstLineIndent: ABNT_CONSTANTS.INDENTATION.FIRST_LINE_PT, // 1.25 cm
      });

      return {
        success: true,
        fontUsed: fontName,
        sectionsModified: marginsResult.count || 1,
        paragraphsFormatted: docResult.count || 0,
        message: `Documento formatado com sucesso! Margens ABNT aplicadas, fonte ${fontName} 12pt, entrelinhas 1.5 e recuo 1,25 cm.`,
      };
    } catch (error) {
      console.error('General ABNT formatting error:', error);
      return {
        success: false,
        error: error.message || 'Falha ao formatar documento.',
      };
    }
  }

  /**
   * Applies only the ABNT margins (without touching typography or spacing)
   */
  static async applyMarginsOnly() {
    try {
      await wordBridge.applyAbntMargins();
      return {
        success: true,
        message: 'Margens ABNT (Superior 3cm, Esquerda 3cm, Inferior 2cm, Direita 2cm) aplicadas com sucesso!',
      };
    } catch (error) {
      return {
        success: false,
        error: error.message,
      };
    }
  }
}
