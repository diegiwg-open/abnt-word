/**
 * Bibliographic Reference Generator – ABNT NBR 6023:2018/2020
 * Formats complete bibliographic entries for books, journal articles,
 * academic works (theses/dissertations), websites, and legislation.
 */

export class ReferenceGenerator {
  /**
   * Formats author names in NBR 6023 style (LAST NAME, First Names)
   * e.g. "SILVA, João; SANTOS, Maria"
   * @param {string} authorsString - Raw author string
   * @returns {string} Formatted author string
   */
  static formatAuthorsNBR6023(authorsString) {
    if (!authorsString || !authorsString.trim()) {
      return 'AUTOR';
    }

    const list = authorsString
      .split(/;| e /i)
      .map((a) => a.trim())
      .filter((a) => a.length > 0);

    const formattedList = list.map((author) => {
      if (author.includes(',')) {
        const parts = author.split(',');
        return `${parts[0].trim().toUpperCase()}, ${parts[1].trim()}`;
      }
      const names = author.split(/\s+/);
      if (names.length === 1) return names[0].toUpperCase();
      const lastName = names[names.length - 1].toUpperCase();
      const firstNames = names.slice(0, names.length - 1).join(' ');
      return `${lastName}, ${firstNames}`;
    });

    if (formattedList.length <= 3) {
      return formattedList.join('; ');
    } else {
      // More than 3 authors → first author followed by et al.
      return `${formattedList[0]} et al.`;
    }
  }

  /**
   * Book reference (NBR 6023)
   * e.g. SILVA, João da. Metodologia científica. 2. ed. São Paulo: Atlas, 2024.
   * @param {Object} data - { authors, title, subtitle, edition, city, publisher, year, pages }
   */
  static book(data) {
    const authors = this.formatAuthorsNBR6023(data.authors);
    const title = (data.title || 'Título da obra').trim();
    const subtitle = data.subtitle && data.subtitle.trim() ? `: ${data.subtitle.trim()}` : '';
    const edition = data.edition && data.edition.trim() ? ` ${data.edition.trim()}. ed.` : '';
    const city = (data.city || 'Local').trim();
    const publisher = (data.publisher || 'Editora').trim();
    const year = (data.year || 'Ano').trim();
    const pages = data.pages && data.pages.trim() ? ` ${data.pages.trim()} p.` : '';

    return `${authors}. **${title}**${subtitle}.${edition} ${city}: ${publisher}, ${year}.${pages}`;
  }

  /**
   * Journal article reference (NBR 6023)
   * e.g. SANTOS, Maria; SILVA, João. Inovações em IA na educação. Revista Brasileira, v. 32, n. 1, p. 10-25, 2024.
   * @param {Object} data - { authors, articleTitle, journalName, city, volume, number, pages, year, doi }
   */
  static journalArticle(data) {
    const authors = this.formatAuthorsNBR6023(data.authors);
    const articleTitle = (data.articleTitle || 'Título do artigo').trim();
    const journalName = (data.journalName || 'Nome do Periódico').trim();
    const city = data.city && data.city.trim() ? ` ${data.city.trim()},` : '';
    const volume = data.volume && data.volume.trim() ? ` v. ${data.volume.trim()},` : '';
    const number = data.number && data.number.trim() ? ` n. ${data.number.trim()},` : '';
    const pages = data.pages && data.pages.trim() ? ` p. ${data.pages.trim()},` : '';
    const year = (data.year || 'Ano').trim();
    const doi = data.doi && data.doi.trim()
      ? ` DOI: https://doi.org/${data.doi.trim().replace(/^https?:\/\/doi\.org\//i, '')}.`
      : '';

    return `${authors}. ${articleTitle}. **${journalName}**,${city}${volume}${number}${pages} ${year}.${doi}`;
  }

  /**
   * Academic work reference – thesis, dissertation, or undergraduate paper (NBR 6023)
   * e.g. OLIVEIRA, Carlos. Título. 2024. Dissertação (Mestrado) – Universidade, Cidade, 2024.
   * @param {Object} data - { authors, title, subtitle, year, yearDefense, degreeType, institution, city, leaves }
   */
  static academicWork(data) {
    const authors = this.formatAuthorsNBR6023(data.authors);
    const title = (data.title || 'Título do trabalho').trim();
    const subtitle = data.subtitle && data.subtitle.trim() ? `: ${data.subtitle.trim()}` : '';
    const yearDefense = (data.yearDefense || data.year || 'Ano').trim();
    const degreeType = (data.degreeType || 'Monografia (Graduação em Engenharia)').trim();
    const institution = (data.institution || 'Universidade').trim();
    const city = (data.city || 'Local').trim();
    const year = (data.year || yearDefense).trim();
    const leaves = data.leaves && data.leaves.trim() ? ` ${data.leaves.trim()} f.` : '';

    return `${authors}. **${title}**${subtitle}. ${yearDefense}.${leaves} ${degreeType} – ${institution}, ${city}, ${year}.`;
  }

  /**
   * Website / online article reference (NBR 6023)
   * e.g. BRASIL. Ministério da Educação. Censo 2023. Brasília, 2024. Disponível em: https://... Acesso em: 15 mar. 2024.
   * @param {Object} data - { authors, title, siteName, year, url, accessDate }
   */
  static website(data) {
    const authors = this.formatAuthorsNBR6023(data.authors || data.organization);
    const title = (data.title || 'Título da página ou matéria').trim();
    const siteName = data.siteName && data.siteName.trim() ? ` **${data.siteName.trim()}**,` : '';
    const year = (data.year || 'Ano').trim();
    const url = (data.url || 'https://...').trim();
    const accessDate = (data.accessDate || 'Acesso em: dia mês. ano').trim();

    return `${authors}. ${title}.${siteName} ${year}. Disponível em: <${url}>. Acesso em: ${accessDate}.`;
  }

  /**
   * Legislation / law reference (NBR 6023)
   * e.g. BRASIL. Lei nº 14.133, de 1º de abril de 2021. [...]. Diário Oficial da União, 1 abr. 2021.
   * @param {Object} data - { jurisdiction, lawName, summary, publication, date }
   */
  static legislation(data) {
    const jurisdiction = (data.jurisdiction || 'BRASIL').trim().toUpperCase();
    const lawName = (data.lawName || 'Lei nº 0.000, de 1 de janeiro de 2024').trim();
    const summary = (data.summary || 'Ementa ou descrição da lei').trim();
    const publication = (data.publication || 'Diário Oficial da União: Brasília, DF').trim();
    const date = (data.date || '1 jan. 2024').trim();

    return `${jurisdiction}. **${lawName}**. ${summary}. ${publication}, ${date}.`;
  }
}
