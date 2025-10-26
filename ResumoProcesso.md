# Técnicas e Recursos Utilizados na Automação
Este documento apresenta de forma resumida as principais técnicas e recursos empregados no processo de automação do projeto DownloadExcel.

---

## Técnicas Utilizadas

- **Automação Web com BotCity**
  - Utilização do framework BotCity WebBot para interação direta com elementos da página web.
  - Seleção e manipulação de elementos DOM via expressões XPath para garantir precisão no preenchimento dos formulários.
  - Controle do navegador Firefox via WebDriver gerenciado pelo GeckoDriverManager, conforme mostrado no curso de introdução a automação Web da BotCity.

- **Estratégia de Tentativas e Tratamento de Exceções**
  - Repetição do processo até 3 vezes em caso de falhas, para garantir robustez.
  - Tratamento de perda de sessão e erros inesperados, com logs detalhados para diagnóstico.
  
- **Manipulação de Dados com Pandas**
  - Utilização da biblioteca pandas para leitura e processamento do arquivo Excel.
  - Limpeza e normalização dos dados (como remoção de espaços extras nas colunas).

- **Controle de Sessão e Navegador**
  - Verificação da sessão ativa para evitar erros durante o processamento.
  - Finalização controlada do navegador para evitar instâncias duplicadas ou travadas.

- **Logging Detalhado**
  - Registro de eventos, erros e fluxos de execução em arquivos de log diários.
  - Facilita monitoramento e auditoria do processo.

---

## Recursos Tecnológicos

- **Python 3** como linguagem principal para o desenvolvimento da automação.
- **BotCity Framework** para automação web robusta via WebBot.
- **GeckoDriverManager** para gerenciamento automático do driver do navegador Firefox.
- **Pandas** para manipulação eficiente de dados tabulares (Excel).
- **Logging padrão do Python** para sistema de logs estruturado.
- **XPath** como técnica para precisão na localização de elementos web.

---

Este conjunto de técnicas e ferramentas oferece um fluxo automatizado confiável, que aborda tanto a interação direta com interfaces web quanto a manipulação de dados, assegurando resiliência e controle na automação do desafio técnico.
