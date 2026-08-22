/**
 * Citation Generator – ABNT NBR 10520:2023
 * Supports the classic uppercase format and the updated 2023 author-date format.
 */

export class CitationGenerator {
  /**
   * Formats author surnames for in-text citations
   * @param {string} authorsString - e.g. "Silva, João; Santos, Maria" or "João Silva"
   * @param {boolean} uppercase - If true → "SILVA; SANTOS", if false → "Silva e Santos"
   */
  static formatAuthors(authorsString, uppercase = true) {
    if (!authorsString || !authorsString.trim()) {
      return uppercase ? 'AUTOR' : 'Autor';
    }

    // Split authors by semicolons or " e "
    const authorList = authorsString
      .split(/;| e /i)
      .map((a) => a.trim())
      .filter((a) => a.length > 0);

    const surnames = authorList.map((full) => {
      // "Silva, João" → take the first part before the comma
      if (full.includes(',')) {
        return full.split(',')[0].trim();
      }
      // "João da Silva" → take the last word as surname
      const parts = full.split(/\s+/);
      return parts[parts.length - 1];
    });

    if (surnames.length === 1) {
      return uppercase ? surnames[0].toUpperCase() : this.capitalize(surnames[0]);
    } else if (surnames.length === 2) {
      if (uppercase) {
        return `${surnames[0].toUpperCase()}; ${surnames[1].toUpperCase()}`;
      } else {
        return `${this.capitalize(surnames[0])} e ${this.capitalize(surnames[1])}`;
      }
    } else if (surnames.length === 3) {
      if (uppercase) {
        return `${surnames[0].toUpperCase()}; ${surnames[1].toUpperCase()}; ${surnames[2].toUpperCase()}`;
      } else {
        return `${this.capitalize(surnames[0])}, ${this.capitalize(surnames[1])} e ${this.capitalize(surnames[2])}`;
      }
    } else {
      // More than 3 authors → et al.
      const first = uppercase ? surnames[0].toUpperCase() : this.capitalize(surnames[0]);
      return `${first} et al.`;
    }
  }

  /**
   * Generates an indirect in-text citation
   * e.g. "Segundo Silva (2024)..." or "De acordo com Silva e Santos (2024)..."
   */
  static generateInText(authors, year, page = '') {
    const authorFormatted = this.formatAuthors(authors, false);
    const pageStr = page && page.trim() ? `, p. ${page.trim()}` : '';
    return `${authorFormatted} (${year}${pageStr})`;
  }

  /**
   * Generates a parenthetical citation (classic author-date system)
   * e.g. "(SILVA, 2024, p. 15)" or "(Silva, 2024, p. 15)"
   */
  static generateParenthetical(authors, year, page = '', uppercase = true) {
    const authorFormatted = this.formatAuthors(authors, uppercase);
    const pageStr = page && page.trim() ? `, p. ${page.trim()}` : '';
    return `(${authorFormatted}, ${year}${pageStr})`;
  }

  /**
   * Generates a short direct quotation (up to 3 lines, inline)
   * e.g. "Quoted text here" (SILVA, 2024, p. 15).
   */
  static generateDirectShort(quoteText, authors, year, page = '', uppercase = true) {
    const ref = this.generateParenthetical(authors, year, page, uppercase);
    return `"${quoteText.trim()}" ${ref}.`;
  }

  /**
   * Generates an apud (secondary source) citation
   * e.g. (SILVA, 1995 apud SANTOS, 2024, p. 50)
   */
  static generateApud(originalAuthor, originalYear, citedAuthor, citedYear, page = '') {
    const orig = this.formatAuthors(originalAuthor, true);
    const cited = this.formatAuthors(citedAuthor, true);
    const pageStr = page && page.trim() ? `, p. ${page.trim()}` : '';
    return `(${orig}, ${originalYear} apud ${cited}, ${citedYear}${pageStr})`;
  }

  /** Capitalizes only the first letter of a string */
  static capitalize(str) {
    if (!str) return '';
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  }
}
