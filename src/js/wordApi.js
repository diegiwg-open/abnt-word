/**
 * Word API Abstraction Layer (Office.js)
 * Compatible with Word 2021, 2024, 365, and Word on the Web.
 * Falls back to WordMock simulator when running outside Microsoft Word (e.g. browser).
 */

import { ABNT_CONSTANTS, cmToPt, ptToCm } from './utils/abntRules.js';
import { mockDocInstance } from './utils/wordMock.js';

export class WordBridge {
  constructor() {
    this.isOfficeAvailable = typeof Office !== 'undefined' && typeof Word !== 'undefined';
  }

  /**
   * Returns true when running inside the real Microsoft Word host
   */
  isRealWord() {
    return typeof Office !== 'undefined' && typeof Word !== 'undefined' && Office.context && Office.context.requirements;
  }

  /**
   * Applies official ABNT margins to all document sections (NBR 14724)
   */
  async applyAbntMargins() {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const sections = context.document.sections;
        context.load(sections, 'body/pageSetup');
        await context.sync();

        for (let i = 0; i < sections.items.length; i++) {
          const pageSetup = sections.items[i].body.pageSetup;
          pageSetup.topMargin = ABNT_CONSTANTS.MARGINS.TOP_PT;
          pageSetup.leftMargin = ABNT_CONSTANTS.MARGINS.LEFT_PT;
          pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
          pageSetup.rightMargin = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
          pageSetup.paperSize = 'A4';
        }

        await context.sync();
        return { success: true, count: sections.items.length };
      });
    } else {
      // Mock fallback
      mockDocInstance.sections.forEach((sec) => {
        sec.pageSetup.topMargin = ABNT_CONSTANTS.MARGINS.TOP_PT;
        sec.pageSetup.leftMargin = ABNT_CONSTANTS.MARGINS.LEFT_PT;
        sec.pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
        sec.pageSetup.rightMargin = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
        sec.pageSetup.paperSize = 'A4';
      });
      return { success: true, count: mockDocInstance.sections.length };
    }
  }

  /**
   * Applies full ABNT formatting to the entire document body
   * @param {Object} options - { fontName, fontSize, lineSpacing, firstLineIndent }
   */
  async formatEntireDocument(options = {}) {
    const fontName = options.fontName || ABNT_CONSTANTS.FONTS.ARIAL;
    const bodyFontSize = options.fontSize || ABNT_CONSTANTS.FONTS.SIZES.BODY;
    const lineSpacing = options.lineSpacing || ABNT_CONSTANTS.LINE_SPACING.BODY; // 1.5
    const firstLineIndent = options.firstLineIndent !== undefined
      ? options.firstLineIndent
      : ABNT_CONSTANTS.INDENTATION.FIRST_LINE_PT;

    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        // Load sections for margin application
        const sections = context.document.sections;
        context.load(sections, 'body/pageSetup');

        // Load all paragraphs
        const paragraphs = context.document.body.paragraphs;
        context.load(paragraphs, 'text,font,lineSpacing,firstLineIndent,leftIndent,alignment,spaceBefore,spaceAfter,style');
        await context.sync();

        // Apply ABNT margins
        for (let i = 0; i < sections.items.length; i++) {
          const pageSetup = sections.items[i].body.pageSetup;
          pageSetup.topMargin = ABNT_CONSTANTS.MARGINS.TOP_PT;
          pageSetup.leftMargin = ABNT_CONSTANTS.MARGINS.LEFT_PT;
          pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
          pageSetup.rightMargin = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
        }

        let formattedCount = 0;

        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          const text = p.text ? p.text.trim() : '';

          if (!text) continue;

          // Detect paragraph type
          const isHeading = (p.style && p.style.toLowerCase().includes('heading')) ||
                            (p.style && p.style.toLowerCase().includes('título')) ||
                            /^\d+(\.\d+)*\s+[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ]/.test(text);

          const isLongQuote = p.leftIndent && p.leftIndent >= 100;

          const isReference = text.includes('NBR ') ||
                              (/^[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ\s]{3,},\s+[A-Z]/.test(text) && text.length > 50);

          // Set font
          if (p.font) {
            p.font.name = fontName;
            p.font.color = ABNT_CONSTANTS.FONTS.COLOR_BLACK;
          }

          if (isHeading) {
            if (p.font) p.font.size = bodyFontSize;
            p.alignment = 'Left';
            p.firstLineIndent = 0;
            p.lineSpacing = lineSpacing * 12;
          } else if (isLongQuote) {
            if (p.font) p.font.size = ABNT_CONSTANTS.FONTS.SIZES.LONG_CITATION;
            p.leftIndent = ABNT_CONSTANTS.INDENTATION.LONG_CITATION_LEFT_PT;
            p.firstLineIndent = 0;
            p.lineSpacing = 12; // single
            p.alignment = 'Justified';
            p.spaceBefore = 6;
            p.spaceAfter = 6;
          } else if (isReference) {
            if (p.font) p.font.size = bodyFontSize;
            p.leftIndent = 0;
            p.firstLineIndent = 0;
            p.lineSpacing = 12; // single
            p.alignment = 'Left';
            p.spaceBefore = 0;
            p.spaceAfter = 6;
          } else {
            // Standard body text
            if (p.font) p.font.size = bodyFontSize;
            p.lineSpacing = lineSpacing * 12; // 1.5x → 18pt
            p.alignment = 'Justified';
            p.leftIndent = 0;
            p.firstLineIndent = firstLineIndent; // 1.25 cm
            p.spaceBefore = 0;
            p.spaceAfter = 0;
          }

          formattedCount++;
        }

        await context.sync();
        return { success: true, count: formattedCount };
      });
    } else {
      // Mock fallback
      mockDocInstance.sections.forEach((sec) => {
        sec.pageSetup.topMargin = ABNT_CONSTANTS.MARGINS.TOP_PT;
        sec.pageSetup.leftMargin = ABNT_CONSTANTS.MARGINS.LEFT_PT;
        sec.pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
        sec.pageSetup.rightMargin = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
      });

      mockDocInstance.paragraphs.forEach((p) => {
        p.font.name = fontName;
        p.font.color = ABNT_CONSTANTS.FONTS.COLOR_BLACK;

        if (p.isHeading) {
          p.font.size = 12;
          p.alignment = 'Left';
          p.firstLineIndent = 0;
          p.lineSpacing = 1.5;
        } else {
          p.font.size = 12;
          p.alignment = 'Justified';
          p.firstLineIndent = 35.43;
          p.lineSpacing = 1.5;
          p.spaceBefore = 0;
          p.spaceAfter = 0;
        }
      });

      return { success: true, count: mockDocInstance.paragraphs.length };
    }
  }

  /**
   * Applies formatting to the current selection (active paragraph or selection range)
   * @param {Object} formatting - Formatting properties to apply
   */
  async formatSelection(formatting) {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        context.load(paragraphs, 'text,font,lineSpacing,firstLineIndent,leftIndent,rightIndent,alignment,spaceBefore,spaceAfter');
        await context.sync();

        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];

          if (p.font) {
            if (formatting.fontName) p.font.name = formatting.fontName;
            if (formatting.fontSize) p.font.size = formatting.fontSize;
            if (formatting.bold !== undefined) p.font.bold = formatting.bold;
            if (formatting.italic !== undefined) p.font.italic = formatting.italic;
            if (formatting.color) p.font.color = formatting.color;
          }

          if (formatting.alignment) p.alignment = formatting.alignment;
          if (formatting.lineSpacing !== undefined) {
            p.lineSpacing = formatting.lineSpacing * 12;
          }
          if (formatting.firstLineIndent !== undefined) p.firstLineIndent = formatting.firstLineIndent;
          if (formatting.leftIndent !== undefined) p.leftIndent = formatting.leftIndent;
          if (formatting.rightIndent !== undefined) p.rightIndent = formatting.rightIndent;
          if (formatting.spaceBefore !== undefined) p.spaceBefore = formatting.spaceBefore;
          if (formatting.spaceAfter !== undefined) p.spaceAfter = formatting.spaceAfter;
        }

        await context.sync();
        return { success: true, count: paragraphs.items.length };
      });
    } else {
      // Mock fallback
      const p = mockDocInstance.paragraphs[mockDocInstance.selectedParagraphIndex];
      if (p) {
        if (formatting.fontName) p.font.name = formatting.fontName;
        if (formatting.fontSize) p.font.size = formatting.fontSize;
        if (formatting.bold !== undefined) p.font.bold = formatting.bold;
        if (formatting.italic !== undefined) p.font.italic = formatting.italic;
        if (formatting.alignment) p.alignment = formatting.alignment;
        if (formatting.lineSpacing) p.lineSpacing = formatting.lineSpacing;
        if (formatting.firstLineIndent !== undefined) p.firstLineIndent = formatting.firstLineIndent;
        if (formatting.leftIndent !== undefined) p.leftIndent = formatting.leftIndent;
        if (formatting.spaceBefore !== undefined) p.spaceBefore = formatting.spaceBefore;
        if (formatting.spaceAfter !== undefined) p.spaceAfter = formatting.spaceAfter;
      }
      return { success: true, count: 1 };
    }
  }

  /**
   * Inserts text at the cursor position or at the end of the document
   * @param {string} text - Text to insert
   * @param {'Selection'|'End'} location - Insertion point
   */
  async insertText(text, location = 'End') {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        let range;
        if (location === 'Selection') {
          range = context.document.getSelection();
        } else {
          range = context.document.body;
        }

        range.insertParagraph(text, location === 'Selection' ? 'After' : 'End');
        await context.sync();
        return { success: true };
      });
    } else {
      mockDocInstance.paragraphs.push({
        text,
        font: { name: 'Arial', size: 12, bold: false, italic: false, color: '#000000' },
        alignment: 'Justified',
        lineSpacing: 1.5,
        firstLineIndent: 35.43,
        spaceBefore: 0,
        spaceAfter: 0,
        isHeading: false,
      });
      return { success: true };
    }
  }

  /**
   * Scans the entire document and returns structured data for the ABNT auditor
   */
  async scanDocument() {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const sections = context.document.sections;
        context.load(sections, 'body/pageSetup');

        const paragraphs = context.document.body.paragraphs;
        context.load(paragraphs, 'text,font,lineSpacing,firstLineIndent,leftIndent,alignment,spaceBefore,spaceAfter,style');
        await context.sync();

        const sectionData = sections.items.map((sec) => ({
          topMargin: sec.body.pageSetup.topMargin,
          leftMargin: sec.body.pageSetup.leftMargin,
          bottomMargin: sec.body.pageSetup.bottomMargin,
          rightMargin: sec.body.pageSetup.rightMargin,
          paperSize: sec.body.pageSetup.paperSize,
        }));

        const paragraphData = paragraphs.items.map((p) => ({
          text: p.text || '',
          fontName: (p.font && p.font.name) || 'Arial',
          fontSize: (p.font && p.font.size) || 12,
          bold: (p.font && p.font.bold) || false,
          italic: (p.font && p.font.italic) || false,
          color: (p.font && p.font.color) || '#000000',
          lineSpacing: p.lineSpacing || 18,
          firstLineIndent: p.firstLineIndent || 0,
          leftIndent: p.leftIndent || 0,
          alignment: p.alignment || 'Justified',
          spaceBefore: p.spaceBefore || 0,
          spaceAfter: p.spaceAfter || 0,
          style: p.style || 'Normal',
        }));

        return { sections: sectionData, paragraphs: paragraphData };
      });
    } else {
      // Mock fallback
      return {
        sections: mockDocInstance.sections.map((s) => ({ ...s.pageSetup })),
        paragraphs: mockDocInstance.paragraphs.map((p) => ({
          text: p.text,
          fontName: p.font.name,
          fontSize: p.font.size,
          bold: p.font.bold,
          italic: p.font.italic,
          color: p.font.color,
          lineSpacing: p.lineSpacing ? p.lineSpacing * 12 : 18,
          firstLineIndent: p.firstLineIndent,
          leftIndent: p.leftIndent || 0,
          alignment: p.alignment,
          spaceBefore: p.spaceBefore,
          spaceAfter: p.spaceAfter,
          style: p.isHeading ? `Heading ${p.headingLevel || 1}` : 'Normal',
        })),
      };
    }
  }

  /**
   * Sorts the selected reference paragraphs alphabetically (A-Z, pt-BR locale)
   */
  async sortReferences() {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        context.load(paragraphs, 'text');
        await context.sync();

        if (paragraphs.items.length <= 1) {
          return { success: false, message: 'Selecione dois ou mais parágrafos de referências.' };
        }

        const items = paragraphs.items
          .map((p) => p.text.trim())
          .filter((t) => t.length > 0);

        items.sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));

        for (let i = 0; i < paragraphs.items.length && i < items.length; i++) {
          paragraphs.items[i].insertText(items[i], 'Replace');
          paragraphs.items[i].alignment = 'Left';
          paragraphs.items[i].firstLineIndent = 0;
          paragraphs.items[i].leftIndent = 0;
          paragraphs.items[i].lineSpacing = 12;
          paragraphs.items[i].spaceAfter = 6;
        }

        await context.sync();
        return { success: true, count: items.length };
      });
    } else {
      return { success: true, count: 2 };
    }
  }

  /**
   * Removes duplicate whitespace from the document body
   */
  async cleanDocument() {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const doubleSpaces = body.search('  ', { matchCase: false });
        context.load(doubleSpaces);
        await context.sync();

        let spaceFixCount = 0;
        for (let i = 0; i < doubleSpaces.items.length; i++) {
          doubleSpaces.items[i].insertText(' ', 'Replace');
          spaceFixCount++;
        }

        await context.sync();
        return { success: true, spacesFixed: spaceFixCount };
      });
    } else {
      return { success: true, spacesFixed: 5 };
    }
  }
}

export const wordBridge = new WordBridge();
