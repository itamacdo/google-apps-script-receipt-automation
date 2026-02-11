# Gerador Automático de Recibos – Google Apps Script

## Visão Geral
Este projeto automatiza o fluxo de faturamento, transformando dados de uma planilha do Google Sheets em recibos formatados no Google Docs. A solução é integrada ao Google Workspace, gerenciando automaticamente a criação de documentos, organização de pastas no Drive e inserção de logotipos dinâmicos.

A arquitetura foi pensada para ser **modular**, permitindo que diferentes empresas ou entidades tenham layouts e dados bancários específicos dentro do mesmo script.

---

## Funcionalidades
* **Interface Customizada:** Menu personalizado no Google Sheets para acesso rápido às funções.
* **Filtro Inteligente:** Diálogo interativo (HTML/JavaScript) que permite gerar recibos por **Número da Nota** ou **Data do Protocolo**.
* **Logística de Pastas:** Criação automática de hierarquia de pastas no Drive baseada no nome da planilha e na entidade (Empresa A/Empresa B).
* **Prevenção de Duplicidade:** O script verifica se o arquivo já existe antes de criar um novo, evitando redundâncias.
* **Processamento de Texto:** Conversão automática de valores numéricos para **extenso** (reais e centavos).
* **Layouts Dinâmicos:** Inserção de logotipos redimensionados proporcionalmente e formatação de cabeçalhos e assinaturas.

---

## Estrutura de Pastas Esperada no Drive
Para o pleno funcionamento, o script busca e organiza os arquivos seguindo este padrão:
`ARQUIVOS_GERAIS` > `RECIBOS_GERADOS` > `[Nome da Planilha]` > `[Subpastas das Empresas]`

---

## Tecnologias Utilizadas
* **Google Apps Script:** Motor principal da automação.
* **JavaScript (ES6+):** Lógica de filtragem, manipulação de arrays e conversão de valores.
* **HTML/CSS:** Para a interface do usuário (modais de busca).
* **Google Sheets API:** Extração e manipulação de dados da planilha.
* **Google Docs API:** Geração e formatação de documentos profissionais.
* **Google Drive API:** Gestão de arquivos, pastas e permissões.

---

## Lógica da Solução
1.  **Entrada de Dados:** O usuário aciona o menu e define os critérios de busca (Nota ou Data) via modal HTML.
2.  **Filtragem:** O script varre a planilha ativa, identifica a qual entidade o dado pertence e valida se os campos obrigatórios estão preenchidos.
3.  **Gestão de Arquivos:** Localiza ou cria as pastas de destino de forma hierárquica no Drive.
4.  **Construção do Documento:** * Cria um novo Google Doc baseado no conteúdo da linha.
    * Insere o logotipo correspondente à empresa.
    * Aplica o layout específico (Dados bancários, CNPJ e cabeçalho).
    * Formata o corpo do texto com o valor formatado e por extenso.
5.  **Finalização:** Move o arquivo da raiz para a pasta final e oferece um link direto para a pasta de saída.

---

## Nota de Segurança e Privacidade
> **IMPORTANTE:** Todos os dados sensíveis (Nomes de Instituições, CNPJs, Contas Bancárias e Localizações) foram removidos ou substituídos por placeholders (`EMPRESA_A`, `00.000.000/0000-00`, etc.) nesta versão pública. O objetivo deste repositório é demonstrar a **lógica de programação e automação de processos**, garantindo a conformidade com as melhores práticas de proteção de dados e privacidade.

---

## Resultados Obtidos
* **Eficiência Operacional:** Redução drástica no tempo de emissão manual de documentos.
* **Eliminação de Erros:** Mitigação de falhas na digitação de valores por extenso e dados bancários.
* **Padronização:** Garantia de que todos os documentos seguem a identidade visual e legal correta de forma automática.

---
*Projeto desenvolvido para fins de automação de processos e demonstração técnica em portfólio profissional.*
