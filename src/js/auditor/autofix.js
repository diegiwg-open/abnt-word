/**
 * ABNT Auto-Fix Engine
 * Executes targeted or batch corrections for issues detected by the Auditor.
 */

import { GeneralFormatter } from '../formatters/generalFormatter.js';
import { CleanerFormatter } from '../formatters/cleanerFormatter.js';
import { wordBridge } from '../wordApi.js';
import { ABNT_CONSTANTS } from '../utils/abntRules.js';

export class AbntAutoFix {
  /**
   * Runs the specified fix action
   * @param {string} actionId - Action identifier (e.g. 'fix_margins', 'fix_fonts', 'fix_all')
   * @param {Object} options - Optional parameters (e.g. { fontName })
   */
  static async executeFix(actionId, options = {}) {
    const fontName = options.fontName || ABNT_CONSTANTS.FONTS.ARIAL;

    switch (actionId) {
      case 'fix_margins':
        return await GeneralFormatter.applyMarginsOnly();

      case 'fix_fonts':
      case 'fix_spacing':
      case 'fix_indents':
      case 'fix_alignment':
      case 'fix_all':
        return await GeneralFormatter.formatFullDocument({ fontName });

      case 'clean_spaces':
        return await CleanerFormatter.cleanDocument();

      case 'fix_heading_dots':
        return await this.fixHeadingDots();

      default:
        return { success: false, message: 'Ação de correção desconhecida.' };
    }
  }

  /**
   * Removes erroneous dots after section numbers in heading paragraphs
   * e.g. "1. INTRODUÇÃO" → "1 INTRODUÇÃO"
   */
  static async fixHeadingDots() {
    if (wordBridge.isRealWord()) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        context.load(paragraphs, 'text');
        await context.sync();

        let count = 0;
        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          const text = p.text ? p.text.trim() : '';

          if (/^\d+(\.\d+)*\.\s+/.test(text)) {
            const newText = text.replace(/^(\d+(\.\d+)*)\.\ +/, '$1 ');
            p.insertText(newText, 'Replace');
            count++;
          }
        }

        await context.sync();
        return { success: true, count, message: `${count} título(s) corrigido(s)!` };
      });
    } else {
      return { success: true, count: 1, message: 'Títulos corrigidos no simulador.' };
    }
  }
}
