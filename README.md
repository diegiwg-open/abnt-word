# 📑 Formatador ABNT Automático para Microsoft Word (2021 / 2024 / 365)

Extensão e Suplemento oficial desenvolvido para o **Microsoft Word 2021**, **Word 2024**, **Word 365** e **Word na Web** para automatizar, formatar, auditar e gerar trabalhos acadêmicos (TCC, Monografias, Artigos Científicos, Dissertações e Teses) rigorosamente dentro das normas da **ABNT**.

---

## ⚡ Como Usar (1 Único Script Idempotente)

Basta executar o script mestre **`ABNT.ps1`** (ou dar duplo-clique no **`ABNT.bat`**). 

Ele verifica dependências, certificados SSL, inicia o servidor local em segundo plano e abre o Microsoft Word com a aba **"Normas ABNT"** carregada automaticamente.

### 1. Criar um Novo Documento:
```powershell
.\ABNT.ps1
```

### 2. Abrir um Documento Existente com o Suplemento:
```powershell
.\ABNT.ps1 "caminho\para\seu_trabalho.docx"
```
*(No Windows Explorer, você também pode simplesmente arrastar qualquer arquivo `.docx` para cima do arquivo `ABNT.bat`!)*

---

## 🌟 Funcionalidades Principais

### 🚀 1. Formatação Geral com 1-Clique
- **Margens Oficiais (NBR 14724):**
  - Superior: 3,0 cm | Esquerda: 3,0 cm
  - Inferior: 2,0 cm | Direita: 2,0 cm
- **Tipografia e Parágrafos:**
  - Seleção entre **Arial** ou **Times New Roman**
  - Corpo de texto em 12 pt, cor preta padronizada
  - Espaçamento entre linhas de 1,5
  - Recuo de primeira linha de 1,25 cm (1 tabulação)
  - Alinhamento Justificado e espaçamentos 0 pt antes/depois
  - Configuração automática do papel em formato A4 (21 x 29,7 cm)

---

### 🔍 2. Auditor ABNT em Tempo Real (Linter & Diagnosticador)
- Analisa todo o documento em tempo real e calcula a **Nota de Conformidade ABNT (0 a 100%)**.
- Detecta instantaneamente:
  - Margens incorretas
  - Múltiplas fontes misturadas ou fontes inválidas
  - Parágrafos sem recuo de 1,25 cm
  - Espaçamentos incorretos (ex: entrelinhas simples ou duplo no corpo do texto)
  - Títulos com ponto indevido após o número (ex: `1. INTRODUÇÃO` ➔ `1 INTRODUÇÃO`)
  - Espaços em branco duplos e excesso de linhas vazias
- Botão **"⚡ Corrigir Todos os Erros"** para autofix instantâneo de todas as não conformidades!

---

### 📐 3. Formatador Rápido de Seleção e Títulos (NBR 6024)
Aplique estilos ABNT com um único clique em qualquer parágrafo selecionado:
- **📄 Corpo de Texto:** 12pt, 1.5 entrelinhas, 1.25cm recuo, justificado.
- **❝ Citação Direta Longa (> 3 linhas):** Recuo de 4,0 cm da margem esquerda, tamanho 10pt, entrelinha simples, sem aspas.
- **🏷️ 1 SEÇÃO PRIMÁRIA:** 12pt, Negrito, Todas Maiúsculas (Caixa Alta).
- **📑 1.1 Seção Secundária:** 12pt, Sem Negrito, Todas Maiúsculas.
- **📌 1.1.1 Seção Terciária:** 12pt, Negrito, Apenas 1ª Maiúscula.
- **🔹 1.1.1.1 Seção Quaternária:** 12pt, Normal, Apenas 1ª Maiúscula.
- **🖼️ Legenda de Figura/Tabela:** 10pt, Centralizado, Simples.
- **📝 Fonte / Nota de Rodapé:** 10pt, Alinhado à Esquerda, Simples.
- **📚 Item de Referência:** 12pt, Alinhado à Esquerda, Simples, 6pt após o parágrafo.

---

### 📚 4. Gerador Inteligente de Citações & Referências (NBR 10520:2023 & NBR 6023)
- Formulário intuitivo para gerar referências e citações automáticas:
  - 📘 Livros e Capítulos de Livros
  - 📄 Artigos de Periódicos / Revistas Científicas
  - 🎓 Teses, Dissertações e TCCs
  - 🌐 Websites e Artigos Online
  - ⚖️ Legislação, Leis e Decretos
- Gera automaticamente:
  - Citação no texto: `(SILVA, 2024, p. 45)` e `Silva (2024, p. 45)`
  - Referência bibliográfica completa formatada segundo a NBR 6023
- **🔤 Organizador Alfabético (A-Z):** Selecione sua lista de referências no Word e ordene-as em ordem alfabética estrita com 1 clique!

---

### 📑 5. Inserção de Modelos Estruturados
Insira com 1 clique modelos pré-formatados:
- **📘 Capa ABNT**
- **📄 Folha de Rosto** (com nota de natureza do trabalho recuada a 8 cm à direita)
- **📝 Resumo & Palavras-chave** (parágrafo único de 150 a 500 palavras)
- **🌐 Abstract & Keywords**
- **📋 Modelo de Sumário**
- **🖼️ Modelo de Figura** (com Topo e Base conforme IBGE/ABNT)
- **📊 Modelo de Tabela**

---

### 🧹 6. Higienizador do Documento
- Remove espaços duplos e múltiplos espaços repetidos.
- Corrige pontuações de seções.
- Limpa quebras de parágrafo excessivas.

---

## 🧪 Testes Automatizados

Para rodar a suíte de testes de validação das regras ABNT:
```bash
npm test
```
Resultados: 13 testes cobrindo todas as normas vigentes, conversões métricas, geradores de citações, referências e diagnósticos de auditoria com 100% de aprovação.
