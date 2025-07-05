# AGENTS.md – regras fixas (projeto VB6)

## Escopo
Válido para **todos** os arquivos `.FRM`, `.BAS`, `.CLS`, `.VBP`, `.RC` e scripts BAT deste repositório.

---

## 1. Codificação
- `.FRM` `.BAS` `.CLS` `.RC` `.VBP` → **ANSI (Windows-1252)**
- Demais arquivos (Markdown, Python, PowerShell, etc.) → **UTF-8**

---

## 2. Acentos (obrigatório)
- **Nunca** remover, trocar ou simplificar diacríticos.  
- Exemplos que DEVEM aparecer exatamente como no código-fonte:  
  "Descrição", "Preço", "Sequência", "Seqüência", "Município",  
  "Último", "Função", "Água", "Ródio".
- Se existirem duas grafias ( Sequência × Seqüência ), preserve a forma encontrada — **não padronizar**.
- **Strings SQL** devem permanecer 100 % idênticas; não alterar nem quebrar.  
  Exemplo intocável:  
  ```vb
  strSql = "SELECT Número, Descrição FROM Peças WHERE Série = 'Óleo'"
  ```

---

## Orientações ao agente
- Responder sempre em português.
- Manter intactos todos os acentos, cedilhas e caracteres especiais.
- Ao alterar ou sugerir trechos de código, explicar o motivo de cada modificação.
- Em comentários ou explicações internas, diferenciar claramente texto explicativo do próprio código.
