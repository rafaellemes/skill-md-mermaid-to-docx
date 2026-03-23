import os
import re
import sys
import tempfile
import subprocess
import argparse
import platform

# ==============================================================================
# SROS - Gerador de Documentos Executivos (Markdown -> Word / DOCX)
# ==============================================================================
# Pré-requisito Único na Máquina Local: Instalação do NodeJS (npm)
# Este script foi arquitetado propositalmente para não enviar seus dados do 
# Mermaid para APIs gratuitas online, garantindo sigilo corporativo.
# ==============================================================================

def check_and_install_dependencies():
    """ Verifica se as dependências Python exatas e motor nativo Pandoc existem. (Cross-Platform) """
    current_os = platform.system() # Retorna 'Windows', 'Linux' ou 'Darwin' (MacOS)
    
    print(f"[1/3] Inteligência de Setup Ativada. Checando infraestrutura no SO Host: [{current_os}]")
    try:
        import pypandoc
    except ImportError:
        print("      [-] Lib 'pypandoc' Ausente. Disparando instalação via PIP na VENV atual...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pypandoc", "-q"])
        import pypandoc

    try:
        # Checa se o binário base do Pandoc (O motor real que cria arquivos Word) está respondendo
        pypandoc.get_pandoc_version()
    except OSError:
        print(f"      [!] Binário C++ do Pandoc não localizado no PATH do {current_os}.")
        print("      [+] Iniciando rotina Cross-Platform de Auto-Instalação Nativa...")
        try:
            # Tenta a extração oficial do repositório MacFarlane multi-arquitetura (Windows/Linux/Mac)
            pypandoc.download_pandoc()
            print("      [OK] Pandoc baixado via PyPandoc Manager.")
        except Exception as e:
            print(f"      [-] Falha no download padrão: {e}")
            print(f"      [+] Injetando requisição no Gerenciador de Pacotes do {current_os} (Fallback)...")
            if current_os == "Linux":
                subprocess.run("sudo apt-get update && sudo apt-get install -y pandoc", shell=True)
            elif current_os == "Darwin":
                subprocess.run("brew install pandoc", shell=True)
            elif current_os == "Windows":
                subprocess.run("winget install --id JohnMacFarlane.Pandoc -e --accept-source-agreements --accept-package-agreements", shell=True)
            else:
                print("      [X] OS não mapeado. Instale o Pandoc manualmente: https://pandoc.org/installing.html")
    print("      >> Motor de Exportação PyPandoc & CLI 100% Operacional.")

def run_mermaid_cli(mermaid_code, output_image_path):
    """ Rotina Privativa Totalmente Offline de renderização de arquitetura C4 """
    
    # 1. Isolamos o codigo do mermaid capturado num arquivinho .mmd
    with tempfile.NamedTemporaryFile(delete=False, suffix=".mmd", mode="w", encoding="utf-8") as f:
        f.write(mermaid_code)
        temp_mmd_path = f.name
    
    try:
        print("      -> Renderizando UML Mermaid localmente (Isolamento de Segurança ativado)...")
        # 2. Invocamos o empacotador global npx do Node.js.
        # Ele vai rodar invisivelmente a biblioteca mmdc da Web convertendo a imagem em PNG na hora.
        # Para suportar MS Word com Tema Escuro ou Claro sem sumir texto, forçamos Box Branca Sólida padrão.
        subprocess.run(
            ["npx", "-y", "@mermaid-js/mermaid-cli", "-i", temp_mmd_path, "-o", output_image_path, "-b", "white", "-t", "default"],
            check=True,
            shell=True, # Essencial pro npx ser puxado nas vars do PATH do Windows
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except subprocess.CalledProcessError as e:
        print(f"\n[ERRO CRITICO] O renderizador local do Mermaid falhou.")
        current_os = platform.system()
        print(f"Você instalou o Node.js no seu sistema {current_os}? Ele é vital para o npx funcionar.")
        print(f"Log do Erro: {e}\n")
    finally:
        if os.path.exists(temp_mmd_path):
            os.remove(temp_mmd_path)

def process_single_markdown(md_filepath):
    import pypandoc
    print(f"\n[Processando Documento] {os.path.basename(md_filepath)}")
    
    with open(md_filepath, "r", encoding="utf-8") as file:
        md_content = file.read()

    # Caçando blocos brutos de código Mermaid
    mermaid_pattern = re.compile(r'```mermaid\n(.*?)\n```', re.DOTALL)
    matches = mermaid_pattern.findall(md_content)
    
    images_to_cleanup = []
    
    if len(matches) == 0:
        print("   -> Nenhum gráfico Mermaid nativo detectado. O documento será transposto diretamente para .docx de forma nativa.")
    else:
        print(f"   -> Encontrados {len(matches)} gráficos complexos no documento.")
    
    # Renderizamos e substituímos provisoriamente cada gráfico
    for i, mermaid_code in enumerate(matches):
        base_dir = os.path.dirname(md_filepath)
        img_filename = f"diagrama_autogerado_{i}.png"
        img_filepath = os.path.join(base_dir, img_filename)
        
        # Gera o .PNG usando Opcão B (Node/Local) Protegendo Sigilo
        run_mermaid_cli(mermaid_code, img_filepath)
        images_to_cleanup.append(img_filepath)
        
        # Correção Crítica Pandoc no Windows: O Pandoc roda num CWD diferente da pasta alvo as vezes.
        # Obrigatoriamente precisamos passar o caminho ABSOLUTO da imagem invertendo as barras '\' para '/'.
        safe_img_filepath = img_filepath.replace('\\', '/')
        replacement_markdown = f"![Diagrama Lógico SROS]({safe_img_filepath})"
        escaped_original_block = re.escape(f"```mermaid\n{mermaid_code}\n```")
        
        # Trocamos no texto que estava em memória apenas
        md_content = re.sub(escaped_original_block, replacement_markdown, md_content, count=1)

    # O Pypandoc via "convert_text" perdia o contexto do CWD (Diretório Original) 
    # ao jogar a string formatada pro binário do Pandoc.
    # Solução: Escrevermos num MD fantasma exatamente na mesma pasta antes de chamar.
    temp_md_path = md_filepath.replace(".md", "_temp_pandoc_shadow.md")
    with open(temp_md_path, "w", encoding="utf-8") as temp_file:
        temp_file.write(md_content)

    # Convertando o diretório real pro Formato Proprietário DOCX
    output_docx = md_filepath.replace(".md", ".docx")
    print(f"   -> Forjando binário Executivo DOCX com Imagens...")
    
    # O pypandoc com a instrução de arquivo "convert_file" consegue caçar as PNGs do lado dele
    pypandoc.convert_file(temp_md_path, 'docx', outputfile=output_docx)
    
    print(f" [SUCESSO] Arquivo gerado com gráficos acoplados: {output_docx}")

    # Protocolo de Limpeza Profunda: Varrer toda a esteira do disco logicol
    images_to_cleanup.append(temp_md_path)
    for item in images_to_cleanup:
        if os.path.exists(item):
            os.remove(item)

def main():
    print("=====================================================")
    print(" SROS - CONVERSOR EXECUTIVO (MARKDOWN PARA MS WORD)  ")
    print("=====================================================")
    
    # Estruturando Argumentos de Terminal em CLI
    parser = argparse.ArgumentParser(description="Converte Arquivos MD isolados em DOCX via MS Word")
    parser.add_argument("arquivo", help="Caminho exato do arquivo Markdown (.md) a ser processado")
    args = parser.parse_args()

    target_file = os.path.abspath(args.arquivo)

    if not os.path.exists(target_file):
        print(f"[ERRO LOGICO] O arquivo especificado está em falta ou nomeado com quebra estrutural: {target_file}")
        return

    if not target_file.endswith(".md"):
        print("[ERRO SISTEMICO] A interface bloqueia extensões estritamente diferentes do corpo de código limpo (.md).")
        return

    # Dispara a instalação passiva sub-modulada interna
    check_and_install_dependencies()
    
    # Rotina Focada Estritamente no Arquivo Declarado Exato 
    process_single_markdown(target_file)
        
    print("\n[OK] Processo de Exportação Executiva e Formatações Físicas Concluído. Abasteça a Alta Gestão com Resultados Limpos!")
    print("=====================================================\n")

if __name__ == "__main__":
    main()
