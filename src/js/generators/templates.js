/**
 * Academic Document Templates (ABNT NBR 14724)
 * Provides ready-made structural templates for direct insertion into Word.
 */

import { wordBridge } from '../wordApi.js';

export class TemplateGenerator {
  /**
   * Generates a Cover Page (Capa) template structure
   * @param {string} institution - Institution name
   * @param {string} author - Full author name
   * @param {string} title - Paper title
   * @param {string} subtitle - Paper subtitle (optional)
   * @param {string} city - City and state (e.g. "Brasília – DF")
   * @param {string} year - Publication year
   */
  static getCapaTemplate(institution, author, title, subtitle, city, year) {
    const inst = (institution || 'NOME DA INSTITUIÇÃO DE ENSINO\nFACULDADE OU DEPARTAMENTO').toUpperCase();
    const aut = (author || 'NOME COMPLETO DO AUTOR').toUpperCase();
    const tit = (title || 'TÍTULO DO TRABALHO ACADÊMICO').toUpperCase();
    const sub = subtitle ? `: ${subtitle.toUpperCase()}` : '';
    const cit = (city || 'CIDADE – UF').toUpperCase();
    const yr = year || new Date().getFullYear().toString();

    return `${inst}\n\n\n\n\n\n${aut}\n\n\n\n\n\n${tit}${sub}\n\n\n\n\n\n\n\n\n${cit}\n${yr}`;
  }

  /**
   * Generates a Title Page (Folha de Rosto) template structure
   * @param {string} author - Full author name
   * @param {string} title - Paper title
   * @param {string} subtitle - Paper subtitle (optional)
   * @param {string} natureText - Nature note (e.g. degree, advisor)
   * @param {string} city - City and state
   * @param {string} year - Year
   */
  static getFolhaDeRostoTemplate(author, title, subtitle, natureText, city, year) {
    const aut = (author || 'NOME COMPLETO DO AUTOR').toUpperCase();
    const tit = (title || 'TÍTULO DO TRABALHO ACADÊMICO').toUpperCase();
    const sub = subtitle ? `: ${subtitle.toUpperCase()}` : '';
    const nature = natureText || 'Trabalho de Conclusão de Curso apresentado ao Curso de Graduação como requisito parcial para a obtenção do título de Bacharel em ...\n\nOrientador: Prof. Dr. Nome do Orientador';
    const cit = (city || 'CIDADE – UF').toUpperCase();
    const yr = year || new Date().getFullYear().toString();

    return `${aut}\n\n\n\n\n\n${tit}${sub}\n\n\n\n\n[NOTA_NATUREZA]\n${nature}\n[/NOTA_NATUREZA]\n\n\n\n\n\n${cit}\n${yr}`;
  }

  /**
   * Generates an Abstract (Resumo) and Keywords (Palavras-chave) template (NBR 6028)
   * @param {string} summaryText - Abstract body text
   * @param {string} keywords - Semicolon-separated keywords
   */
  static getResumoTemplate(summaryText, keywords) {
    const text = summaryText || 'O resumo deve ressaltar o objetivo, o método, os resultados e as conclusões do trabalho. O texto deve ser composto por uma sequência de frases concisas, afirmativas e não de enumeração de tópicos. Deve ser redigido em parágrafo único, justificado, sem recuo de primeira linha, com extensão de 150 a 500 palavras para trabalhos de conclusão de curso, monografias, dissertações e teses.';
    const kws = keywords || 'Palavra-chave 1. Palavra-chave 2. Palavra-chave 3.';

    return `RESUMO\n\n${text}\n\nPalavras-chave: ${kws}`;
  }

  /**
   * Generates an English Abstract template
   * @param {string} abstractText - Abstract body in English
   * @param {string} keywords - Keywords in English
   */
  static getAbstractTemplate(abstractText, keywords) {
    const text = abstractText || 'The abstract is the English version of the summary. It must highlight the objective, method, results, and conclusions of the academic work in a single justified paragraph without first-line indentation, consisting of 150 to 500 words.';
    const kws = keywords || 'Keyword 1. Keyword 2. Keyword 3.';

    return `ABSTRACT\n\n${text}\n\nKeywords: ${kws}`;
  }

  /**
   * Generates an ABNT Figure identification template
   * (label above the figure + source attribution below)
   * @param {string} number - Figure number
   * @param {string} title - Descriptive figure title
   * @param {string} source - Source attribution
   */
  static getFiguraTemplate(number, title, source) {
    const num = number || '1';
    const tit = title || 'Título explicativo da ilustração';
    const src = source || 'Elaborado pelo autor (2024).';

    return `Figura ${num} – ${tit}\n\n[INSERIR IMAGEM / GRÁFICO / ILUSTRAÇÃO AQUI]\n\nFonte: ${src}`;
  }

  /**
   * Generates an ABNT / IBGE Table template
   * @param {string} number - Table number
   * @param {string} title - Descriptive table title
   * @param {string} source - Data source attribution
   */
  static getTabelaTemplate(number, title, source) {
    const num = number || '1';
    const tit = title || 'Título explicativo da tabela';
    const src = source || 'Dados da pesquisa (2024).';

    return `Tabela ${num} – ${tit}\n\n| Categoria | Variável A | Variável B | Total |\n| :--- | :---: | :---: | :---: |\n| Item 1 | 10 | 25 | 35 |\n| Item 2 | 15 | 30 | 45 |\n| **Total Geral** | **25** | **55** | **80** |\n\nFonte: ${src}`;
  }

  /**
   * Inserts the specified template into the Word document at the current cursor position
   * @param {string} type - Template type: 'capa', 'folha_rosto', 'resumo', 'abstract', 'figura', 'tabela', 'sumario'
   * @param {Object} data - Optional data to populate the template fields
   */
  static async insertTemplateIntoDocument(type, data = {}) {
    let content = '';

    switch (type) {
      case 'capa':
        content = this.getCapaTemplate(data.institution, data.author, data.title, data.subtitle, data.city, data.year);
        break;
      case 'folha_rosto':
        content = this.getFolhaDeRostoTemplate(data.author, data.title, data.subtitle, data.natureText, data.city, data.year);
        break;
      case 'resumo':
        content = this.getResumoTemplate(data.summaryText, data.keywords);
        break;
      case 'abstract':
        content = this.getAbstractTemplate(data.abstractText, data.keywords);
        break;
      case 'figura':
        content = this.getFiguraTemplate(data.number, data.title, data.source);
        break;
      case 'tabela':
        content = this.getTabelaTemplate(data.number, data.title, data.source);
        break;
      case 'sumario':
        content = `SUMÁRIO\n\n1 INTRODUÇÃO ........................................................................ 10\n2 REVISÃO DA LITERATURA ..................................................... 12\n  2.1 CONCEITOS FUNDAMENTAIS ............................................. 14\n  2.2 TRABALHOS CORRELATOS .............................................. 18\n3 METODOLOGIA ..................................................................... 22\n4 RESULTADOS E DISCUSSÃO ................................................. 28\n5 CONSIDERAÇÕES FINAIS ...................................................... 35\nREFERÊNCIAS ........................................................................ 38`;
        break;
      default:
        content = '';
    }

    return await wordBridge.insertText(content, 'Selection');
  }
}
