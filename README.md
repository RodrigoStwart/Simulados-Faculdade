
# Simulados - Single Page App

Pequeno sistema SPA para gerar simulados a partir de uma planilha XLSX.

## Entregáveis
- `index.html`
- `styles.css`
- `app.js`
- `Simulados Faculdade.xlsx` (exemplo)
- `README.md`

## Execução (local)
1. Coloque todos os arquivos na mesma pasta.
2. Rode um servidor HTTP (fetch não funciona via `file://`).
   - Python 3: `python -m http.server 8000`
   - Node (serve): `npx serve`
3. Acesse `http://localhost:8000` no navegador.

## Formato esperado da planilha (primeira aba)
- Linha 1 = cabeçalho (opcional)
- Coluna A = matéria
- Coluna B = enunciado
- Coluna C = alternativa A
- Coluna D = alternativa B
- Coluna E = alternativa C
- Coluna F = alternativa D
- Coluna G = gabarito (A/B/C/D)

## Funcionalidades
- Carrega `Simulados Faculdade.xlsx` automaticamente via `fetch`.
- Permite upload manual como fallback.
- Seleciona matéria, escolhe número de questões, embaralha alternativas, ativa temporizador.
- Mostra 1 questão por página com navegação.
- Ao finalizar, exibe resultado detalhado e salva histórico no `localStorage`.
- Exporta sessão atual em CSV e PDF (client-side).
- Tema claro e modo escuro.

## Dependências (CDN)
- SheetJS `xlsx` v0.18.5 — leitura XLSX.
- html2pdf v0.10.1 — exportar para PDF.

## Observações de segurança e acessibilidade
- Textos vindos da planilha são escapados antes de exibição (função `escapeHtml`).
- Inputs e botões acessíveis via teclado.
- Tratar gabarito inválido: questões com gabarito fora de A/B/C/D são ignoradas e avisos aparecem no console.

## Como alterar nome do arquivo
- Edite a constante `DEFAULT_XLSX` em `app.js` para apontar para outro nome de arquivo, ou use o upload manual.

## Testes manuais sugeridos
(veja no enunciado original — inclui verificar caracteres especiais, exportar, histórico, dark mode, etc.)
