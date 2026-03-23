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
