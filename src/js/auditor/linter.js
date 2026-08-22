/**
 * ABNT Compliance Auditor / Linter
 * Performs a deep scan of the document and identifies violations of ABNT standards.
 * Calculates a compliance score (0–100%) and generates a detailed report with fix suggestions.
 */

import { ABNT_CONSTANTS, ptToCm, roundTo } from '../utils/abntRules.js';
import { wordBridge } from '../wordApi.js';

export class AbntLinter {
  /**
   * Runs a full audit on the current document
   * @returns {Object} Audit result with score, status, issues, and summary
   */
  static async auditDocument() {
    try {
      const docData = await wordBridge.scanDocument();
      return this.analyzeDocument(docData);
    } catch (error) {
      console.error('Error auditing document:', error);
      return {
        score: 0,
        status: 'error',
        message: 'Não foi possível ler o documento para auditoria: ' + error.message,
        issues: [],
      };
    }
  }

  /**
   * Analyzes the extracted document structure and returns a full audit report
   * @param {Object} docData - { sections, paragraphs }
   */
  static analyzeDocument(docData) {
    const issues = [];
    let penalty = 0;

    const sections = docData.sections || [];
    const paragraphs = docData.paragraphs || [];

    // 1. Margin audit (NBR 14724)
    if (sections.length > 0) {
      const sec = sections[0];
      const topDiff = Math.abs(sec.topMargin - ABNT_CONSTANTS.MARGINS.TOP_PT);
      const leftDiff = Math.abs(sec.leftMargin - ABNT_CONSTANTS.MARGINS.LEFT_PT);
      const bottomDiff = Math.abs(sec.bottomMargin - ABNT_CONSTANTS.MARGINS.BOTTOM_PT);
      const rightDiff = Math.abs(sec.rightMargin - ABNT_CONSTANTS.MARGINS.RIGHT_PT);

      const isMarginWrong = topDiff > 3 || leftDiff > 3 || bottomDiff > 3 || rightDiff > 3;

      if (isMarginWrong) {
        issues.push({
          id: 'margins_incorrect',
          category: 'Margens',
          severity: 'critical',
          title: 'Margens fora do padrão ABNT',
          description: `Atual: Sup: ${roundTo(ptToCm(sec.topMargin), 1)}cm, Esq: ${roundTo(ptToCm(sec.leftMargin), 1)}cm, Inf: ${roundTo(ptToCm(sec.bottomMargin), 1)}cm, Dir: ${roundTo(ptToCm(sec.rightMargin), 1)}cm. O padrão ABNT exige 3cm (Superior e Esquerda) e 2cm (Inferior e Direita).`,
          fixAction: 'fix_margins',
          fixLabel: 'Ajustar Margens (3,3,2,2)',
        });
        penalty += 25;
      }
    }

    // Statistical counters for paragraph analysis
    const fontUsage = {};
    let nonAbntFontCount = 0;
    let wrongLineSpacingCount = 0;
    let wrongIndentCount = 0;
    let nonJustifiedBodyCount = 0;
    let titleWithDotCount = 0;
    let doubleSpaceCount = 0;
    let totalTextParagraphs = 0;

    paragraphs.forEach((p) => {
      const text = (p.text || '').trim();
      if (!text) return; // skip empty paragraphs

      totalTextParagraphs++;

      // Font usage tracking
      const font = p.fontName || 'Unknown';
      fontUsage[font] = (fontUsage[font] || 0) + 1;
      const isAllowedFont = font.includes('Arial') || font.includes('Times');
      if (!isAllowedFont) nonAbntFontCount++;

      // Heading with erroneous dot (e.g. "1. INTRODUÇÃO")
      if (/^\d+(\.\d+)*\.\s+[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ]/.test(text)) {
        titleWithDotCount++;
      }

      // Double whitespace
      if (p.text && p.text.includes('  ')) {
        doubleSpaceCount++;
      }

      // Determine paragraph type to apply correct rules
      const isHeading = (p.style && p.style.toLowerCase().includes('heading')) ||
                        /^\d+(\.\d+)*\s+[A-Z]/.test(text);
      const isQuote = p.leftIndent && p.leftIndent > 100;
      const isRef = /^[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ\s]{3,},\s+[A-Z]/.test(text) && text.length > 50;

      if (!isHeading && !isQuote && !isRef) {
        // Body text: must have 1.5 line spacing (18pt)
        const lineSpacingVal = p.lineSpacing || 18;
        if (lineSpacingVal < 16 || lineSpacingVal > 20) wrongLineSpacingCount++;

        // First-line indent must be ~1.25 cm (35.43 pt)
        const firstLineIndentVal = p.firstLineIndent || 0;
        if (firstLineIndentVal < 25 || firstLineIndentVal > 45) wrongIndentCount++;

        // Alignment must be Justified
        if (p.alignment && p.alignment !== 'Justified') nonJustifiedBodyCount++;
      }
    });

    // 2. Typography / font audit
    const usedFonts = Object.keys(fontUsage);
    if (usedFonts.length > 1) {
      issues.push({
        id: 'mixed_fonts',
        category: 'Tipografia',
        severity: 'warning',
        title: 'Múltiplas fontes detectadas',
        description: `O trabalho contém mais de uma fonte (${usedFonts.join(', ')}). A ABNT exige padronização exclusiva em Arial ou Times New Roman em todo o documento.`,
        fixAction: 'fix_fonts',
        fixLabel: 'Padronizar Fonte',
      });
      penalty += 15;
    } else if (nonAbntFontCount > 0) {
      issues.push({
        id: 'invalid_font',
        category: 'Tipografia',
        severity: 'warning',
        title: 'Fonte fora das normas ABNT',
        description: `Detectada fonte ${usedFonts[0] || 'incompatível'}. A ABNT recomenda estritamente Arial ou Times New Roman.`,
        fixAction: 'fix_fonts',
        fixLabel: 'Alterar para Arial / Times',
      });
      penalty += 15;
    }

    // 3. Line spacing audit
    if (wrongLineSpacingCount > 0) {
      issues.push({
        id: 'wrong_line_spacing',
        category: 'Espaçamento',
        severity: 'warning',
        title: `${wrongLineSpacingCount} parágrafo(s) com espaçamento incorreto`,
        description: 'O corpo do texto deve ter espaçamento de 1,5 linhas. Foram encontrados parágrafos com espaçamento simples ou irregular.',
        fixAction: 'fix_spacing',
        fixLabel: 'Aplicar Espaçamento 1.5',
      });
      penalty += Math.min(20, wrongLineSpacingCount * 5);
    }

    // 4. First-line indent audit
    if (wrongIndentCount > 0) {
      issues.push({
        id: 'wrong_indent',
        category: 'Parágrafos',
        severity: 'warning',
        title: `${wrongIndentCount} parágrafo(s) sem recuo de 1,25 cm`,
        description: 'Todo parágrafo comum deve ter recuo de 1,25 cm (1 tabulação) na primeira linha.',
        fixAction: 'fix_indents',
        fixLabel: 'Aplicar Recuo de 1,25 cm',
      });
      penalty += Math.min(15, wrongIndentCount * 4);
    }

    // 5. Alignment audit
    if (nonJustifiedBodyCount > 0) {
      issues.push({
        id: 'non_justified',
        category: 'Alinhamento',
        severity: 'info',
        title: `${nonJustifiedBodyCount} parágrafo(s) não justificados`,
        description: 'O texto corrente deve ter alinhamento justificado.',
        fixAction: 'fix_alignment',
        fixLabel: 'Justificar Texto',
      });
      penalty += Math.min(10, nonJustifiedBodyCount * 3);
    }

    // 6. Heading numbering audit (NBR 6024)
    if (titleWithDotCount > 0) {
      issues.push({
        id: 'heading_dots',
        category: 'Títulos NBR 6024',
        severity: 'info',
        title: `${titleWithDotCount} título(s) com ponto indevido após o número`,
        description: 'A NBR 6024 proíbe ponto entre o número e o título (ex: use "1 INTRODUÇÃO" e não "1. INTRODUÇÃO").',
        fixAction: 'fix_heading_dots',
        fixLabel: 'Corrigir Numeração de Títulos',
      });
      penalty += 10;
    }

    // 7. Double spaces audit
    if (doubleSpaceCount > 0) {
      issues.push({
        id: 'double_spaces',
        category: 'Digitação',
        severity: 'info',
        title: `${doubleSpaceCount} parágrafo(s) com espaços duplos repetidos`,
        description: 'Foram encontrados múltiplos espaços em branco consecutivos.',
        fixAction: 'clean_spaces',
        fixLabel: 'Remover Espaços Duplos',
      });
      penalty += 5;
    }

    // Calculate final ABNT compliance score (0–100)
    const score = Math.max(0, 100 - penalty);

    return {
      score,
      status: score >= 90 ? 'excellent' : score >= 70 ? 'good' : score >= 50 ? 'warning' : 'critical',
      totalParagraphs: totalTextParagraphs,
      fontUsage,
      issues,
      summary: {
        marginsOk: !issues.some((i) => i.id === 'margins_incorrect'),
        fontsOk: nonAbntFontCount === 0 && usedFonts.length <= 1,
        spacingOk: wrongLineSpacingCount === 0,
        indentsOk: wrongIndentCount === 0,
        headingsOk: titleWithDotCount === 0,
      },
    };
  }
}
