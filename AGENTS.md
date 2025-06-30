# AGENTS

## 1. Diretrizes de Resposta
- **Responder sempre em português**  
  **Motivo:** Garante consistência linguística e evita misturar termos em inglês, tornando as respostas imediatamente compreensíveis para o seu público.

## 2. Regras de Código SQL
### 2.1. Preservação de SELECT
- **Preservar trechos de código SQL, especialmente cláusulas `SELECT`, sem alterações**  
  **Motivo:** Seus `SELECTs` contêm acentos, cedilhas e nomes de campos exatos do banco de dados; qualquer modificação pode causar erro de sintaxe ou divergência de nomes.

### 2.2. Acentos e Caracteres Especiais
- **Manter intactos todos os acentos, cedilhas e caracteres especiais**  
  **Motivo:** O seu banco de dados utiliza nomes de campos com caracteres acentuados e cedilhados; alterá-los compromete a integridade das consultas e gera erros de mapeamento.

## 3. Alterações e Sugestões de Código
- **Explicar o motivo de cada modificação**  
  **Motivo:** Você precisa compreender o “porquê” por trás de cada ajuste para avaliar seu impacto no sistema e replicar o padrão em outras partes.

- **Evitar reformulações automáticas que mudem o formato ou a estrutura básica dos comandos DML/DDL**  
  **Motivo:** Reformulações podem acidentalmente renomear colunas ou reordenar parâmetros, causando falhas na aplicação.

## 4. Comentários e Explicações Internas
- **Distinguir claramente texto explicativo do próprio código**  
  **Motivo:** Para não confundir anotações com trechos que devem ser enviados ao compilador; facilita a revisão humana posterior.

## 5. Compatibilidade de Codificação
- **Verificar a compatibilidade da codificação de caracteres (UTF-8) nos exemplos fornecidos**  
  **Motivo:** Assegura que os acentos sejam corretamente interpretados no ambiente de desenvolvimento e no banco de dados.
