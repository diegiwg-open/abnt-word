const assert = require('assert');

// Test runner com imports dinâmicos ES Modules
async function runAllTests() {
  console.log('====================================================');
  console.log('🧪 INICIANDO SUÍTE DE TESTES DO FORMATADOR ABNT');
  console.log('====================================================\n');

  let passed = 0;
  let failed = 0;

  function test(name, fn) {
    try {
      fn();
      console.log(`  ✓ PASS: ${name}`);
      passed++;
    } catch (err) {
      console.error(`  ✗ FAIL: ${name}`);
      console.error(`    Erro: ${err.message}`);
      failed++;
    }
  }

  // 1. Testes de Constantes e Conversões
  const { ABNT_CONSTANTS, cmToPt, ptToCm, roundTo } = await import('../src/js/utils/abntRules.js');

  test('Conversão de Unidades: 3.0 cm -> ~85.04 pt e volta', () => {
    const pt = cmToPt(3.0);
    assert.strictEqual(roundTo(pt, 2), 85.04);
    assert.strictEqual(roundTo(ptToCm(pt), 1), 3.0);
  });

  test('Constantes de Margem ABNT NBR 14724', () => {
    assert.strictEqual(ABNT_CONSTANTS.MARGINS.TOP_CM, 3.0);
    assert.strictEqual(ABNT_CONSTANTS.MARGINS.LEFT_CM, 3.0);
    assert.strictEqual(ABNT_CONSTANTS.MARGINS.BOTTOM_CM, 2.0);
    assert.strictEqual(ABNT_CONSTANTS.MARGINS.RIGHT_CM, 2.0);
  });

  // 2. Testes de Citações NBR 10520:2023
  const { CitationGenerator } = await import('../src/js/generators/citations.js');

  test('Citação 1 autor: (SILVA, 2024, p. 15)', () => {
    const res = CitationGenerator.generateParenthetical('João da Silva', '2024', '15', true);
    assert.strictEqual(res, '(SILVA, 2024, p. 15)');
  });

  test('Citação 2 autores: (SILVA; SANTOS, 2024)', () => {
    const res = CitationGenerator.generateParenthetical('Silva, João; Santos, Maria', '2024', '', true);
    assert.strictEqual(res, '(SILVA; SANTOS, 2024)');
  });

  test('Citação > 3 autores: (SILVA et al., 2024)', () => {
    const res = CitationGenerator.generateParenthetical('Silva, J.; Santos, M.; Oliveira, C.; Costa, F.', '2024', '', true);
    assert.strictEqual(res, '(SILVA et al., 2024)');
  });

  test('Citação no texto: Silva (2024, p. 10)', () => {
    const res = CitationGenerator.generateInText('João da Silva', '2024', '10');
    assert.strictEqual(res, 'Silva (2024, p. 10)');
  });

  test('Citação Apud: (SILVA, 1990 apud SANTOS, 2024, p. 5)', () => {
    const res = CitationGenerator.generateApud('Silva, João', '1990', 'Santos, Maria', '2024', '5');
    assert.strictEqual(res, '(SILVA, 1990 apud SANTOS, 2024, p. 5)');
  });

  // 3. Testes de Referências NBR 6023
  const { ReferenceGenerator } = await import('../src/js/generators/references.js');

  test('Referência de Livro NBR 6023', () => {
    const ref = ReferenceGenerator.book({
      authors: 'João da Silva',
      title: 'Metodologia Científica',
      edition: '3',
      city: 'São Paulo',
      publisher: 'Atlas',
      year: '2024',
    });
    assert(ref.includes('SILVA, João da.'));
    assert(ref.includes('**Metodologia Científica**'));
    assert(ref.includes('3. ed. São Paulo: Atlas, 2024.'));
  });

  test('Referência de Artigo de Revista NBR 6023', () => {
    const ref = ReferenceGenerator.journalArticle({
      authors: 'Maria Santos',
      articleTitle: 'Inteligência Artificial na Saúde',
      journalName: 'Revista Médica',
      city: 'Rio de Janeiro',
      volume: '15',
      number: '2',
      pages: '100-115',
      year: '2024',
    });
    assert(ref.includes('SANTOS, Maria.'));
    assert(ref.includes('**Revista Médica**'));
    assert(ref.includes('v. 15, n. 2, p. 100-115, 2024.'));
  });

  test('Referência de Website NBR 6023', () => {
    const ref = ReferenceGenerator.website({
      authors: 'BRASIL. Ministério da Educação',
      title: 'Diretrizes Curriculares',
      siteName: 'Portal MEC',
      year: '2024',
      url: 'https://www.gov.br/mec',
      accessDate: '15 mar. 2024',
    });
    assert(ref.includes('Disponível em: <https://www.gov.br/mec>'));
    assert(ref.includes('Acesso em: 15 mar. 2024.'));
  });

  // 4. Testes do Auditor ABNT (Linter)
  const { AbntLinter } = await import('../src/js/auditor/linter.js');

  test('Auditor ABNT: detecta documento não conforme e penaliza', () => {
    const mockDoc = {
      sections: [
        { topMargin: 72, leftMargin: 72, bottomMargin: 72, rightMargin: 72 }, // 2.54 cm (não é 3,3,2,2)
      ],
      paragraphs: [
        { text: '1. INTRODUÇÃO', fontName: 'Calibri', fontSize: 16, lineSpacing: 12, firstLineIndent: 0, alignment: 'Left' },
        { text: 'Texto de exemplo sem recuo e em Calibri.', fontName: 'Calibri', fontSize: 11, lineSpacing: 14, firstLineIndent: 0, alignment: 'Left' },
      ],
    };

    const audit = AbntLinter.analyzeDocument(mockDoc);
    assert(audit.score < 80, `Score deveria ser menor que 80, foi: ${audit.score}`);
    assert(audit.issues.some((i) => i.id === 'margins_incorrect'), 'Deveria detectar margens incorretas');
    assert(audit.issues.some((i) => i.id === 'invalid_font'), 'Deveria detectar fonte inválida');
    assert(audit.issues.some((i) => i.id === 'heading_dots'), 'Deveria detectar ponto indevido no título');
  });

  test('Auditor ABNT: documento 100% conforme recebe score 100', () => {
    const perfectDoc = {
      sections: [
        {
          topMargin: ABNT_CONSTANTS.MARGINS.TOP_PT,
          leftMargin: ABNT_CONSTANTS.MARGINS.LEFT_PT,
          bottomMargin: ABNT_CONSTANTS.MARGINS.BOTTOM_PT,
          rightMargin: ABNT_CONSTANTS.MARGINS.RIGHT_PT,
        },
      ],
      paragraphs: [
        { text: '1 INTRODUÇÃO', fontName: 'Arial', fontSize: 12, lineSpacing: 18, firstLineIndent: 0, alignment: 'Left', style: 'Heading 1' },
        { text: 'Texto devidamente formatado conforme todas as normas.', fontName: 'Arial', fontSize: 12, lineSpacing: 18, firstLineIndent: 35.43, alignment: 'Justified' },
      ],
    };

    const audit = AbntLinter.analyzeDocument(perfectDoc);
    assert.strictEqual(audit.score, 100, `Score esperado: 100, obtido: ${audit.score}`);
    assert.strictEqual(audit.issues.length, 0);
  });

  // 5. Testes de Modelos Pré-Textuais
  const { TemplateGenerator } = await import('../src/js/generators/templates.js');

  test('Template de Capa ABNT contém partes obrigatórias', () => {
    const capa = TemplateGenerator.getCapaTemplate('Universidade de Brasília', 'João Silva', 'TCC Inteligência Artificial', '', 'Brasília', '2024');
    assert(capa.includes('UNIVERSIDADE DE BRASÍLIA'));
    assert(capa.includes('JOÃO SILVA'));
    assert(capa.includes('TCC INTELIGÊNCIA ARTIFICIAL'));
    assert(capa.includes('BRASÍLIA'));
  });

  console.log('\n====================================================');
  console.log(`📊 RESULTADOS DOS TESTES: ${passed} PASSOU | ${failed} FALHOU`);
  console.log('====================================================\n');

  if (failed > 0) {
    process.exit(1);
  }
}

runAllTests().catch((err) => {
  console.error('Erro fatal ao rodar testes:', err);
  process.exit(1);
});
