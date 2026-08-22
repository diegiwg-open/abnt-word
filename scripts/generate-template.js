const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  PageBreak,
  convertInchesToTwip,
  convertMillimetersToTwip,
} = require('docx');

async function generateDocxTemplate() {
  const templatesDir = path.join(__dirname, '..', 'templates');
  if (!fs.existsSync(templatesDir)) {
    fs.mkdirSync(templatesDir, { recursive: true });
  }

  // Margens ABNT em Twips: 30mm = ~1701 twips, 20mm = ~1134 twips
  const topMargin = convertMillimetersToTwip(30);
  const leftMargin = convertMillimetersToTwip(30);
  const bottomMargin = convertMillimetersToTwip(20);
  const rightMargin = convertMillimetersToTwip(20);

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: topMargin,
              left: leftMargin,
              bottom: bottomMargin,
              right: rightMargin,
            },
          },
        },
        children: [
          // ==================== CAPA ====================
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, after: 1200 },
            children: [
              new TextRun({
                text: 'NOME DA INSTITUIÇÃO DE ENSINO SUPERIOR',
                font: 'Arial',
                size: 24, // 12 pt
                bold: true,
              }),
              new TextRun({
                text: '\nFACULDADE OU DEPARTAMENTO ACADÊMICO',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, after: 1400 },
            children: [
              new TextRun({
                text: 'NOME COMPLETO DO AUTOR',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, after: 2000 },
            children: [
              new TextRun({
                text: 'TÍTULO DO TRABALHO: SUBTÍTULO CONFORME NORMAS ABNT',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360 },
            children: [
              new TextRun({
                text: 'CIDADE – UF\n2024',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({ children: [new PageBreak()] }),

          // ==================== FOLHA DE ROSTO ====================
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, after: 1200 },
            children: [
              new TextRun({
                text: 'NOME COMPLETO DO AUTOR',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, after: 1000 },
            children: [
              new TextRun({
                text: 'TÍTULO DO TRABALHO: SUBTÍTULO',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          // Nota de aprovação com recuo à esquerda de 8cm (~4535 twips)
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            indent: { left: convertMillimetersToTwip(80) },
            spacing: { line: 240, after: 1600 }, // Simples
            children: [
              new TextRun({
                text: 'Trabalho de Conclusão de Curso apresentado ao Colegiado do Curso de Graduação como requisito parcial para a obtenção do título de Bacharel em ...\n\nOrientador: Prof. Dr. Nome do Orientador',
                font: 'Arial',
                size: 20, // 10 pt
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360 },
            children: [
              new TextRun({
                text: 'CIDADE – UF\n2024',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({ children: [new PageBreak()] }),

          // ==================== RESUMO ====================
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, after: 360 },
            children: [
              new TextRun({
                text: 'RESUMO',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { line: 360, after: 240 },
            children: [
              new TextRun({
                text: 'O resumo deve apresentar os pontos relevantes do texto de forma clara, concisa e objetiva. Composto por um parágrafo único, sem recuo na primeira linha e espaçamento de 1,5 entre linhas (ou simples conforme a instituição), contendo de 150 a 500 palavras. Deve destacar os objetivos, a metodologia utilizada, os principais resultados alcançados e as conclusões do trabalho acadêmico.',
                font: 'Arial',
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { line: 360, after: 400 },
            children: [
              new TextRun({
                text: 'Palavras-chave: ',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
              new TextRun({
                text: 'Normas ABNT. Trabalhos Acadêmicos. Formatação Automatizada.',
                font: 'Arial',
                size: 24,
              }),
            ],
          }),

          new Paragraph({ children: [new PageBreak()] }),

          // ==================== 1 INTRODUÇÃO ====================
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { line: 360, before: 240, after: 240 },
            children: [
              new TextRun({
                text: '1 INTRODUÇÃO',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            indent: { firstLine: convertMillimetersToTwip(12.5) }, // 1.25 cm
            spacing: { line: 360, after: 0 },
            children: [
              new TextRun({
                text: 'A elaboração de monografias, dissertações e teses requer a estrita observância das diretrizes técnicas estabelecidas pela ABNT. Este modelo automatizado foi estruturado para fornecer as margens padronizadas (3 cm superior e esquerda, 2 cm inferior e direita), além da tipografia oficial de 12 pontos em Arial ou Times New Roman com entrelinhas de 1,5 linhas.',
                font: 'Arial',
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            indent: { firstLine: convertMillimetersToTwip(12.5) },
            spacing: { line: 360, after: 240 },
            children: [
              new TextRun({
                text: 'A seguir demonstra-se o padrão para citações diretas longas com mais de três linhas, que devem receber recuo de 4 centímetros da margem esquerda, tamanho de fonte 10 pontos e espaçamento simples entre linhas, sem a utilização de aspas.',
                font: 'Arial',
                size: 24,
              }),
            ],
          }),

          // Citação Direta Longa (> 3 linhas)
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            indent: { left: convertMillimetersToTwip(40) }, // 4.0 cm
            spacing: { line: 240, before: 120, after: 120 }, // Simples
            children: [
              new TextRun({
                text: 'As citações diretas, no texto, com mais de três linhas, devem ser destacadas com recuo de 4 cm da margem esquerda, com letra menor que a do texto utilizado e sem as aspas. Recomenda-se espaçamento simples entre as linhas e espaçamento de 6 pt antes e depois do bloco (ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS, 2023, p. 12).',
                font: 'Arial',
                size: 20, // 10 pt
              }),
            ],
          }),

          // ==================== 2 REFERÊNCIAS ====================
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { line: 360, before: 360, after: 240 },
            children: [
              new TextRun({
                text: 'REFERÊNCIAS',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { line: 240, after: 120 }, // Simples, 6pt depois
            children: [
              new TextRun({
                text: 'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ',
                font: 'Arial',
                size: 24,
              }),
              new TextRun({
                text: 'NBR 14724',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
              new TextRun({
                text: ': Informação e documentação — Trabalhos acadêmicos — Apresentação. Rio de Janeiro: ABNT, 2011.',
                font: 'Arial',
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { line: 240, after: 120 },
            children: [
              new TextRun({
                text: 'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ',
                font: 'Arial',
                size: 24,
              }),
              new TextRun({
                text: 'NBR 6023',
                font: 'Arial',
                size: 24,
                bold: true,
              }),
              new TextRun({
                text: ': Informação e documentação — Referências — Elaboração. Rio de Janeiro: ABNT, 2020.',
                font: 'Arial',
                size: 24,
              }),
            ],
          }),
        ],
      },
    ],
  });

  const outputPath = path.join(templatesDir, 'Modelo_ABNT_Automatizado.docx');
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`✅ Modelo Word ABNT gerado com sucesso em: ${outputPath}`);
}

generateDocxTemplate().catch(console.error);
