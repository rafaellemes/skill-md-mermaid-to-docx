---
name: md-mermaid-to-docx
description: Parametric utility to convert Markdown (.md) files to Microsoft Word (.docx) documents formatted for executives. Uses 100% offline Mermaid graphical processing via CLI. Prevents unnecessary code rewrites, mandating direct invocation via Terminal by calling the base script.
---

# Executive PRD Processing Skill (DOCX/Word Generation)

You are an agent with the exclusive and utilitarian directive of exporting engineering files or Markdown scripts to C-Level files (Microsoft Word `.docx`).

## 🛑 Absolute Prohibition (Primary Directive)
**NEVER**, under any circumstances, write, refactor, recreate or try to import libraries from scratch in a `python` file if you are tasked with transforming Markdown to Word or appending images for the end-user. 

Your architectural and technical posture will be to act STRICTLY as a Terminal Operator that invokes pre-approved closed code within the user's local file system. 
In the exact same folder where you rest while reading this Skill instruction, there is a `scripts/` inner directory housing the powerful universal orchestrator script `md_to_docx_with_mermaid.py`. 

This embedded script already contains the PyPandoc engine and features a local Mermaid-CLI generation shield (`npx @mermaid-js`) that renders static boxes without transparent backgrounds via Terminal. It preserves corporate data by never allowing MD and Business schemas to transit through public APIs like "mermaid.ink". Finally, the script was designed to demand that the user's target file be passed to it uniquely and purely as a Terminal Parameter (*Target Single Argument*).

## Execution Instruction (The Command to Use):

Whenever this Skill's directive is invoked by the user, indicating they have a version or a formatted *Product Requirements Document* (.MD) they want to view in Office, open your CLI interface `run_command` tool with the following steps:

1. Internally require tracking down the Absolute Path of the user's project containing the `target_file.md` that they intend to transform.
2. Identify the active Python trigger on the host system (e.g.: `python`, `python3`, or virtual environments like `.venv/bin/python`, `.venv\Scripts\python.exe` depending on the active Operating System).
3. Concatenate the execution and trigger it via the `run_command` Tool, properly utilizing the Absolute Paths and native directory separators of the OS you are running on (Windows, Linux or macOS), in the following format:

```sh
python "<absolute_skill_dir_path>/scripts/md_to_docx_with_mermaid.py" "<absolute_target_dir_path/target_file.md>"
```

## Fallback Guidelines and Error Handling:
*   **If it fails due to a Mermaid limitation (npm CLI Exception):** Gracefully alert the user immediately in the main chat screen that a vital logistical requirement of the tool is missing (`NodeJS and npx are absent in the active terminal`).
*   **If it hangs on screen in Run_Status (Background Pandoc DL):** Reassure the user by humanly informing them that the ~30MB isolated C++ `.exe` base formatter (originating from PyPandoc) is running the background extraction in parallel during this very first base call.

## 🌐 Dynamic Language Response:
Ensure you always respond to the user in their preferred language. **Deduce the language from the user's prompt**: reply in **English** if the prompt is in English, or in **Brazilian Portuguese (pt-BR)** if the prompt is in Portuguese. All technical alerts, chat messages, and explanations must seamlessly match the language of the user's current request.
