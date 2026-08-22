/**
 * Document Cleaner / Sanitizer
 * Fixes common typing mistakes in academic papers:
 * - Multiple consecutive spaces ("  " → " ")
 * - Multiple empty paragraph breaks
 */

import { wordBridge } from '../wordApi.js';

export class CleanerFormatter {
  /**
   * Runs a general typing cleanup pass on the document
   */
  static async cleanDocument() {
    try {
      const result = await wordBridge.cleanDocument();
      return {
        success: true,
        spacesFixed: result.spacesFixed || 0,
        message: 'Documento higienizado! Espaços duplicados e quebras de linha excessivas corrigidos.',
      };
    } catch (error) {
      return {
        success: false,
        error: error.message,
      };
    }
  }
}
