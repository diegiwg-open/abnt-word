/**
 * Word Simulator (WordMock)
 * Allows the taskpane UI and formatters to work interactively
 * when opened in the browser (outside Microsoft Word) for testing and preview.
 */

export class WordMockDocument {
  constructor() {
    this.sections = [
      {
        pageSetup: {
          topMargin: 72,    // 2.54 cm – intentionally non-ABNT to trigger auditor
          bottomMargin: 72,
          leftMargin: 72,
          rightMargin: 72,
          orientation: 'Portrait',
          paperSize: 'A4',
        },
      },
    ];

    this.paragraphs = [
      {
        text: 'UNIVERSIDADE EXEMPLO DE BRASÍLIA',
        font: { name: 'Calibri', size: 14, bold: true, italic: false, color: '#000000' },
        alignment: 'Centered',
        lineSpacing: 1.15,
        firstLineIndent: 0,
        spaceBefore: 0,
        spaceAfter: 12,
        isHeading: false,
      },
      {
        text: '1 INTRODUÇÃO',
        font: { name: 'Calibri', size: 16, bold: true, italic: false, color: '#000000' },
        alignment: 'Left',
        lineSpacing: 1.15,
        firstLineIndent: 0,
        spaceBefore: 18,
        spaceAfter: 6,
        isHeading: true,
        headingLevel: 1,
      },
      {
        text: 'A elaboração de trabalhos acadêmicos no Brasil exige o cumprimento rigoroso das diretrizes estabelecidas pela Associação Brasileira de Normas Técnicas (ABNT). Este parágrafo de exemplo serve para demonstrar a formatação automatizada de corpo de texto, recuos de primeira linha e entrelinhas.',
        font: { name: 'Calibri', size: 11, bold: false, italic: false, color: '#333333' },
        alignment: 'Left',
        lineSpacing: 1.15,
        firstLineIndent: 0,
        spaceBefore: 0,
        spaceAfter: 8,
        isHeading: false,
      },
      {
        text: 'Conforme preconiza a NBR 10520, as citações diretas no texto com mais de três linhas devem ser destacadas com recuo de quatro centímetros da margem esquerda, com fonte menor que a do texto utilizado e sem aspas.',
        font: { name: 'Arial', size: 12, bold: false, italic: false, color: '#000000' },
        alignment: 'Left',
        lineSpacing: 1.5,
        firstLineIndent: 35.43,
        spaceBefore: 0,
        spaceAfter: 0,
        isHeading: false,
      },
      {
        text: '2 REFERÊNCIAS',
        font: { name: 'Calibri', size: 16, bold: true, italic: false, color: '#000000' },
        alignment: 'Left',
        lineSpacing: 1.15,
        firstLineIndent: 0,
        spaceBefore: 18,
        spaceAfter: 6,
        isHeading: true,
        headingLevel: 1,
      },
      {
        text: 'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. NBR 14724: Informação e documentação — Trabalhos acadêmicos — Apresentação. Rio de Janeiro: ABNT, 2011.',
        font: { name: 'Calibri', size: 11, bold: false, italic: false, color: '#000000' },
        alignment: 'Justified',
        lineSpacing: 1.5,
        firstLineIndent: 35.43,
        spaceBefore: 0,
        spaceAfter: 0,
        isHeading: false,
      },
    ];

    // Index of the paragraph currently "selected" in mock mode
    this.selectedParagraphIndex = 2;
  }
}

export const mockDocInstance = new WordMockDocument();
