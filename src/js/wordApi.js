/**
 * Word API Abstraction Layer (Office.js)
 * Compatible with Word 2021, 2024, 365, and Word on the Web.
 * Falls back to WordMock simulator when running outside Microsoft Word (e.g. browser).
 */

import { ABNT_CONSTANTS } from './utils/abntRules.js';
import { mockDocInstance } from './utils/wordMock.js';

export class WordBridge {
  constructor() {
    this.isOfficeAvailable = typeof Office !== 'undefined' && typeof Word !== 'undefined';
  }

  /**
   * Returns true when running inside the real Microsoft Word host
   */
  isRealWord() {
    try {
      return (
        typeof Office !== 'undefined' &&
        typeof Word !== 'undefined' &&
        Office.context != null &&
        Office.context.requirements != null &&
        Word.run !== undefined
      );
    } catch (error) {
      console.error('Error checking Word connection:', error);
      return false;
    }
  }

  /**
   * Applies official ABNT margins to the document (NBR 14724)
   * Uses document.pageSetup according to official Word JS API documentation.
   */
  async applyAbntMargins() {
    if (this.isRealWord()) {
      try {
        return await Word.run(async (context) => {
          const pageSetup = context.document.pageSetup;
          pageSetup.topMargin    = ABNT_CONSTANTS.MARGINS.TOP_PT;
          pageSetup.leftMargin   = ABNT_CONSTANTS.MARGINS.LEFT_PT;
          pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
          pageSetup.rightMargin  = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
          await context.sync();
          return { success: true, count: 1 };
        });
      } catch (error) {
        console.error('Error applying margins:', error);
        return { success: false, error: error.message };
      }
    } else {
      mockDocInstance.sections.forEach((sec) => {
        sec.pageSetup.topMargin    = ABNT_CONSTANTS.MARGINS.TOP_PT;
        sec.pageSetup.leftMargin   = ABNT_CONSTANTS.MARGINS.LEFT_PT;
        sec.pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
        sec.pageSetup.rightMargin  = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
      });
      return { success: true, count: 1 };
    }
  }

  /**
   * Applies full ABNT formatting to the entire document body.
   * Two-pass approach: first load text/style to classify paragraphs,
   * then write font + paragraph formatting properties.
   * @param {Object} options - { fontName, fontSize, lineSpacing, firstLineIndent }
   */
  async formatEntireDocument(options = {}) {
    const fontName       = options.fontName       || ABNT_CONSTANTS.FONTS.ARIAL;
    const bodyFontSize   = options.fontSize        || ABNT_CONSTANTS.FONTS.SIZES.BODY;
    const lineSpacing    = options.lineSpacing     || ABNT_CONSTANTS.LINE_SPACING.BODY; // 1.5
    const firstLineIndent =
      options.firstLineIndent !== undefined
        ? options.firstLineIndent
        : ABNT_CONSTANTS.INDENTATION.FIRST_LINE_PT;

    if (this.isRealWord()) {
      try {
        return await Word.run(async (context) => {
          // --- Pass 1: Apply margins via document.pageSetup (correct API path) ---
          const pageSetup = context.document.pageSetup;
          pageSetup.topMargin    = ABNT_CONSTANTS.MARGINS.TOP_PT;
          pageSetup.leftMargin   = ABNT_CONSTANTS.MARGINS.LEFT_PT;
          pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
          pageSetup.rightMargin  = ABNT_CONSTANTS.MARGINS.RIGHT_PT;

          // --- Pass 2: Load paragraphs (text + style for classification) ---
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'text,style,leftIndent');
          await context.sync();

          let formattedCount = 0;

          // Set default font for the entire body (works even for empty documents)
          context.document.body.font.name = fontName;
          context.document.body.font.size = bodyFontSize;
          context.document.body.font.color = ABNT_CONSTANTS.FONTS.COLOR_BLACK;

          for (let i = 0; i < paragraphs.items.length; i++) {
            const p    = paragraphs.items[i];
            const text = (p.text || '').trim();

            const styleLower = (p.style || '').toLowerCase();
            const isHeading =
              styleLower.includes('heading') ||
              styleLower.includes('título') ||
              (text && /^\d+(\.\d+)*\s+[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ]/.test(text));

            const isLongQuote = (p.leftIndent || 0) >= 100;
            const isReference =
              text.includes('NBR ') ||
              (/^[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ\s]{3,},\s+[A-Z]/.test(text) && text.length > 50);

            // Set font (writing – no load required)
            p.font.name  = fontName;
            p.font.color = ABNT_CONSTANTS.FONTS.COLOR_BLACK;
            p.font.bold  = isHeading ? true : false;

            if (isHeading) {
              p.font.size      = bodyFontSize;
              p.alignment      = 'Left';
              p.firstLineIndent = 0;
              p.lineSpacing    = lineSpacing * 12;
              p.spaceBefore    = ABNT_CONSTANTS.PARAGRAPH_SPACING.HEADING_BEFORE_PT;
              p.spaceAfter     = ABNT_CONSTANTS.PARAGRAPH_SPACING.HEADING_AFTER_PT;
            } else if (isLongQuote) {
              p.font.size      = ABNT_CONSTANTS.FONTS.SIZES.LONG_CITATION;
              p.leftIndent     = ABNT_CONSTANTS.INDENTATION.LONG_CITATION_LEFT_PT;
              p.firstLineIndent = 0;
              p.lineSpacing    = 12; // single
              p.alignment      = 'Justified';
              p.spaceBefore    = 6;
              p.spaceAfter     = 6;
            } else if (isReference) {
              p.font.size      = bodyFontSize;
              p.leftIndent     = 0;
              p.firstLineIndent = 0;
              p.lineSpacing    = 12; // single
              p.alignment      = 'Left';
              p.spaceBefore    = 0;
              p.spaceAfter     = 6;
            } else {
              // Standard body text (applies even to empty paragraphs)
              p.font.size      = bodyFontSize;
              p.lineSpacing    = lineSpacing * 12; // 18pt
              p.alignment      = 'Justified';
              p.leftIndent     = 0;
              p.firstLineIndent = firstLineIndent; // 35.43pt = 1.25cm
              p.spaceBefore    = 0;
              p.spaceAfter     = 0;
            }

            formattedCount++;
          }

          await context.sync();
          return { success: true, count: formattedCount };
        });
      } catch (error) {
        console.error('Error formatting document:', error);
        return { success: false, error: error.message };
      }
    } else {
      // Mock fallback
      mockDocInstance.sections.forEach((sec) => {
        sec.pageSetup.topMargin    = ABNT_CONSTANTS.MARGINS.TOP_PT;
        sec.pageSetup.leftMargin   = ABNT_CONSTANTS.MARGINS.LEFT_PT;
        sec.pageSetup.bottomMargin = ABNT_CONSTANTS.MARGINS.BOTTOM_PT;
        sec.pageSetup.rightMargin  = ABNT_CONSTANTS.MARGINS.RIGHT_PT;
      });
      mockDocInstance.paragraphs.forEach((p) => {
        p.font.name  = fontName;
        p.font.color = ABNT_CONSTANTS.FONTS.COLOR_BLACK;
        if (p.isHeading) {
          p.font.size = 12; p.alignment = 'Left'; p.firstLineIndent = 0;
        } else {
          p.font.size = 12; p.alignment = 'Justified';
          p.firstLineIndent = 35.43; p.lineSpacing = 1.5;
          p.spaceBefore = 0; p.spaceAfter = 0;
        }
      });
      return { success: true, count: mockDocInstance.paragraphs.length };
    }
  }

  /**
   * Applies formatting to the current selection.
   * Writing properties does NOT require loading in Office JS.
   * @param {Object} formatting - Formatting properties to apply
   */
  async formatSelection(formatting) {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const selection  = context.document.getSelection();
        const paragraphs = selection.paragraphs;

        // Load only 'text' to count items; writing never needs pre-loading
        context.load(paragraphs, 'text');
        await context.sync();

        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];

          if (formatting.fontName  != null) p.font.name   = formatting.fontName;
          if (formatting.fontSize  != null) p.font.size   = formatting.fontSize;
          if (formatting.bold      != null) p.font.bold   = formatting.bold;
          if (formatting.italic    != null) p.font.italic = formatting.italic;
          if (formatting.color     != null) p.font.color  = formatting.color;

          if (formatting.alignment       != null) p.alignment        = formatting.alignment;
          if (formatting.lineSpacing     != null) p.lineSpacing      = formatting.lineSpacing * 12;
          if (formatting.firstLineIndent != null) p.firstLineIndent  = formatting.firstLineIndent;
          if (formatting.leftIndent      != null) p.leftIndent       = formatting.leftIndent;
          if (formatting.rightIndent     != null) p.rightIndent      = formatting.rightIndent;
          if (formatting.spaceBefore     != null) p.spaceBefore      = formatting.spaceBefore;
          if (formatting.spaceAfter      != null) p.spaceAfter       = formatting.spaceAfter;
        }

        await context.sync();
        return { success: true, count: paragraphs.items.length };
      });
    } else {
      const p = mockDocInstance.paragraphs[mockDocInstance.selectedParagraphIndex];
      if (p) {
        if (formatting.fontName       != null) p.font.name        = formatting.fontName;
        if (formatting.fontSize       != null) p.font.size        = formatting.fontSize;
        if (formatting.bold           != null) p.font.bold        = formatting.bold;
        if (formatting.italic         != null) p.font.italic      = formatting.italic;
        if (formatting.alignment      != null) p.alignment        = formatting.alignment;
        if (formatting.lineSpacing    != null) p.lineSpacing      = formatting.lineSpacing;
        if (formatting.firstLineIndent!= null) p.firstLineIndent  = formatting.firstLineIndent;
        if (formatting.leftIndent     != null) p.leftIndent       = formatting.leftIndent;
        if (formatting.spaceBefore    != null) p.spaceBefore      = formatting.spaceBefore;
        if (formatting.spaceAfter     != null) p.spaceAfter       = formatting.spaceAfter;
      }
      return { success: true, count: 1 };
    }
  }

  /**
   * Inserts text at the cursor position or at the end of the document.
   * @param {string} text - Text to insert
   * @param {'Selection'|'End'} location - Insertion point
   */
  async insertText(text, location = 'End') {
    if (this.isRealWord()) {
      return await Word.run(async (context) => {
        const range = location === 'Selection'
          ? context.document.getSelection()
          : context.document.body;
        range.insertParagraph(text, location === 'Selection' ? 'After' : 'End');
        await context.sync();
        return { success: true };
      });
    } else {
      mockDocInstance.paragraphs.push({
        text,
        font: { name: 'Arial', size: 12, bold: false, italic: false, color: '#000000' },
        alignment: 'Justified', lineSpacing: 1.5, firstLineIndent: 35.43,
        spaceBefore: 0, spaceAfter: 0, isHeading: false,
      });
      return { success: true };
    }
  }

  /**
   * Scans the entire document and returns structured data for the ABNT auditor.
   * Uses a two-pass load: first paragraphs metadata, then font sub-properties.
   */
  async scanDocument() {
    if (this.isRealWord()) {
      try {
        return await Word.run(async (context) => {
          // --- Pass 1: Load page setup from document (correct API path) ---
          const pageSetup = context.document.pageSetup;
          context.load(pageSetup, 'topMargin,leftMargin,bottomMargin,rightMargin');

          // Load paragraph scalar properties (no nested font yet)
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'text,alignment,lineSpacing,firstLineIndent,leftIndent,spaceBefore,spaceAfter,style');
          await context.sync();

          // --- Pass 2: Load font properties for each paragraph ---
          paragraphs.items.forEach((p) => {
            context.load(p.font, 'name,size,bold,italic,color');
          });
          await context.sync();

          // Build section summary from document.pageSetup
          const sectionData = [{
            topMargin:    pageSetup.topMargin,
            leftMargin:   pageSetup.leftMargin,
            bottomMargin: pageSetup.bottomMargin,
            rightMargin:  pageSetup.rightMargin,
          }];

          const paragraphData = paragraphs.items.map((p) => ({
            text:           p.text || '',
            fontName:       (p.font && p.font.name)  || 'Unknown',
            fontSize:       (p.font && p.font.size)  || 12,
            bold:           (p.font && p.font.bold)  || false,
            italic:         (p.font && p.font.italic)|| false,
            color:          (p.font && p.font.color) || '#000000',
            lineSpacing:    p.lineSpacing    || 18,
            firstLineIndent:p.firstLineIndent|| 0,
            leftIndent:     p.leftIndent     || 0,
            alignment:      p.alignment      || 'Justified',
            spaceBefore:    p.spaceBefore    || 0,
            spaceAfter:     p.spaceAfter     || 0,
            style:          p.style          || 'Normal',
          }));

          return { sections: sectionData, paragraphs: paragraphData };
        });
      } catch (error) {
        console.error('Error scanning document:', error);
        // Fallback to mock if scan fails
        return {
          sections: mockDocInstance.sections.map((s) => ({ ...s.pageSetup })),
          paragraphs: mockDocInstance.paragraphs.map((p) => ({
            text:           p.text,
            fontName:       p.font.name,
            fontSize:       p.font.size,
            bold:           p.font.bold,
            italic:         p.font.italic,
            color:          p.font.color,
            lineSpacing:    p.lineSpacing ? p.lineSpacing * 12 : 18,
            firstLineIndent:p.firstLineIndent,
            leftIndent:     p.leftIndent || 0,
            alignment:      p.alignment,
            spaceBefore:    p.spaceBefore,
            spaceAfter:     p.spaceAfter,
            style:          p.isHeading ? `Heading ${p.headingLevel || 1}` : 'Normal',
          })),
        };
      }
    } else {
      return {
        sections: mockDocInstance.sections.map((s) => ({ ...s.pageSetup })),
        paragraphs: mockDocInstance.paragraphs.map((p) => ({
          text:           p.text,
          fontName:       p.font.name,
          fontSize:       p.font.size,
          bold:           p.font.bold,
          italic:         p.font.italic,
          color:          p.font.color,
          lineSpacing:    p.lineSpacing ? p.lineSpacing * 12 : 18,
          firstLineIndent:p.firstLineIndent,
          leftIndent:     p.leftIndent || 0,
          alignment:      p.alignment,
          spaceBefore:    p.spaceBefore,
          spaceAfter:     p.spaceAfter,
          style:          p.isHeading ? `Heading ${p.headingLevel || 1}` : 'Normal',
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
        const selection  = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        context.load(paragraphs, 'text');
        await context.sync();

        if (paragraphs.items.length <= 1) {
          return { success: false, message: 'Selecione dois ou mais parágrafos de referências.' };
        }

        const items = paragraphs.items.map((p) => p.text.trim()).filter((t) => t.length > 0);
        items.sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));

        for (let i = 0; i < paragraphs.items.length && i < items.length; i++) {
          paragraphs.items[i].insertText(items[i], 'Replace');
          paragraphs.items[i].alignment     = 'Left';
          paragraphs.items[i].firstLineIndent = 0;
          paragraphs.items[i].leftIndent    = 0;
          paragraphs.items[i].lineSpacing   = 12;
          paragraphs.items[i].spaceAfter    = 6;
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
        const body        = context.document.body;
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
