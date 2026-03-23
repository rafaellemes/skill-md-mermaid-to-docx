# Skill: md-mermaid-to-docx

🇺🇸 **English** | 🇧🇷 **Português**
--- | ---
[English Version](#english) | [Versão em Português](#portugues)

---

<a id="english"></a>
## 🇺🇸 English

### What it does
This repository provides an agentic Skill that converts Markdown (`.md`) files containing Mermaid diagrams into executive-ready Microsoft Word (`.docx`) documents. This utility is 100% offline, executing locally via the CLI to preserve corporate data and avoid transiting schemas over public APIs. It leverages the robust PyPandoc engine and a secure local Mermaid-CLI render, ensuring that diagrams are generated cleanly without transparent backgrounds.

### What is required/installed
When using this Skill, your environment needs the following dependencies:
- **Node.js**: Required to install and run the `npx @mermaid-js/mermaid-cli` natively.
- **Python 3**: To run the backend orchestration script `md_to_docx_with_mermaid.py`.
- **Pandoc**: The underlying C++ binary (`~30MB`) which PyPandoc downloads and utilizes under the hood to format the DOCX file.

### How it works
The agent identifies the target `.md` file in your workspace and programmatically invokes the CLI command invoking the inner `scripts/md_to_docx_with_mermaid.py` with the absolute path of the target file. The Python script then handles the translation, diagram rendering via Mermaid, and produces the finalized Word document inside the target directory.

### Installation
To install this agentic Skill, download or clone this repository and place the `md-mermaid-to-docx` folder within your AI agent's skills directory.

**For Antigravity Projects:**
Antigravity features native auto-discovery for structured skills.
- **Global Installation:** Place the folder in `~/.agents/skills/` (accessible across all projects).
- **Local Installation:** Place the folder in `{.agents, .agent, _agents, _agent}/skills/` within your specific project space.
The Antigravity engine will automatically parse the `SKILL.md` file and incorporate it into its capabilities.

**For Claude Code Projects:**
Claude Code currently does not feature native auto-discovery for skill folders like Antigravity does. To utilize this tool:
- You must supply the directives mapped in `SKILL.md` directly into your Custom Instructions or context.
- *Or*, wrap the `scripts/md_to_docx_with_mermaid.py` execution inside an MCP (Model Context Protocol) Server for Claude Code to access it as a specialized tool.

---

<a id="portugues"></a>
## 🇧🇷 Português (Brasil)

### O que esta Skill faz
Este repositório fornece uma *Skill* agêntica que converte arquivos Markdown (`.md`) contendo diagramas Mermaid em documentos Microsoft Word (`.docx`) com formatação executiva. Este utilitário funciona de maneira 100% offline, executando localmente via CLI para preservar dados corporativos e evitar o tráfego de informações confidenciais por APIs públicas. Ele utiliza o robusto motor PyPandoc juntamente com uma blindagem local para a renderização via Mermaid-CLI, garantindo que os diagramas sejam gerados de maneira limpa (sem fundo transparente).

### O que é necessário/instalado
Ao utilizar esta Skill, seu ambiente precisa das seguintes dependências:
- **Node.js**: Necessário para instalar e rodar nativamente o pacote `npx @mermaid-js/mermaid-cli`.
- **Python 3**: Para rodar o script orquestrador de backend `md_to_docx_with_mermaid.py` (ou em um ambiente virtual).
- **Pandoc**: O binário base em C++ (`~30MB`) que a biblioteca PyPandoc baixa e utiliza nos bastidores na primeira vez para formatar o arquivo base DOCX.

### Como funciona
O agente identifica o arquivo `.md` alvo no seu ambiente local e invoca programaticamente via linha de comando o script interno `scripts/md_to_docx_with_mermaid.py` repassando o caminho absoluto (*absolute path*) do arquivo alvo. O script Python em seguida orquestra toda a tradução, a renderização dos diagramas via Mermaid e compila o documento final do Word no diretório alvo original escolhido.

### Instalação
Para instalar essa Skill, faça o download ou clone este repositório e coloque a pasta `md-mermaid-to-docx` dentro do diretório de *skills* do seu agente de IA.

**Para Projetos Antigravity:**
O Antigravity possui suporte nativo à auto-descoberta dessas instruções.
- **Instalação Global:** Coloque em `~/.agents/skills/` (disponível para todos os seus projetos do SO).
- **Instalação Local:** Coloque em `{.agents, .agent, _agents, _agent}/skills/` dentro da raiz do seu projeto específico.
O motor do Antigravity processará automaticamente o arquivo `SKILL.md` e passará a incorporar essa ferramenta dinamicamente.

**Para Projetos Claude Code:**
O Claude Code não suporta nativamente a leitura de arquivos em pastas pré-determinadas como o Antigravity. Para usá-lo com este modelo:
- Forneça manualmente as diretrizes estritas transcritas no arquivo `SKILL.md` nas suas *Custom Instructions* ou contexto global.
- *Ou*, encapsule a execução de `scripts/md_to_docx_with_mermaid.py` e exponha-o por meio de um Servidor MCP (Model Context Protocol) local para que a IA consiga acionar via MCP Client.
