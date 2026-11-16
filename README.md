# planilhas-design

Guia de estilo visual padrão para planilhas (Excel e Google Sheets) — cores da marca
(`#12424D` petróleo, `#FFB745` âmbar, branco), fontes **Sora** (títulos) e **Poppins** (texto),
alinhamentos, formatos de número e gráficos. Serve para que toda planilha saia sempre com o mesmo estilo,
seja feita por uma pessoa ou por qualquer IA.

## Arquivos

| Arquivo | Para quê |
|---|---|
| `SKILL.md` | O guia completo. É o arquivo principal — funciona como "skill" e como instrução para qualquer IA |
| `tokens.json` | Os mesmos valores (cores, fontes, tamanhos, formatos) em formato de dados |
| `scripts/estilo_excel.py` | Aplica o estilo em arquivos `.xlsx` (Python + openpyxl) |
| `scripts/estilo_sheets.gs` | Aplica o estilo no Google Sheets (Apps Script) |
| `assets/logo.png` | Logo Saldus para fundo escuro (adicionar) |
| `assets/logo-escura.png` | Logo Saldus para fundo branco (adicionar, se existir) |

## Como usar em cada plataforma

**Baixar:** `git clone https://github.com/JuniorCarlini/planilhas-design.git` (ou *Code → Download ZIP* no GitHub).

**Claude Code** — clone/copie a pasta para `~/.claude/skills/planilhas-design/` (vale para todos os projetos)
ou para `.claude/skills/planilhas-design/` dentro de um projeto:
```bash
git clone https://github.com/JuniorCarlini/planilhas-design.git ~/.claude/skills/planilhas-design
```

**Claude.ai (app/web)** — compacte a pasta em `.zip` e envie em *Configurações → Capacidades → Skills*.

**ChatGPT** — crie um GPT (ou Projeto) e cole o conteúdo de `SKILL.md` nas instruções; anexe
`tokens.json` e os scripts como arquivos de conhecimento.

**Gemini** — crie um Gem e cole o conteúdo de `SKILL.md` nas instruções.

**Outras ferramentas que leem SKILL.md / AGENTS.md** (Codex, Cursor etc.) — copie a pasta para a
pasta de skills da ferramenta, ou cole o `SKILL.md` no arquivo de instruções do projeto.

**Sem IA** — use `SKILL.md` como manual e a seção 8 (checklist) para conferir.

## Google Sheets em 1 minuto
1. Na planilha: **Extensões → Apps Script**, cole `scripts/estilo_sheets.gs` e salve.
2. Recarregue a planilha: aparece o menu **Estilo** → *Aplicar estilo padrão*.
3. Logo: **Inserir → Imagem → Imagem sobre as células**, no canto superior direito (linhas 1–2).

## Excel com Python
```bash
pip install openpyxl pillow
python scripts/estilo_excel.py exemplo.xlsx   # gera uma planilha de exemplo no estilo
```
Para o Excel mostrar Sora e Poppins, instale as fontes no computador (grátis em fonts.google.com).
