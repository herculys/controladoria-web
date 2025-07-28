# Controladoria Web Application

Esta aplicação web foi desenvolvida para automatizar o processamento de dados de controladoria, gerando relatórios de produtividade com análises visuais através de gráficos de pizza. A aplicação é construída com Flask e oferece uma interface amigável para upload de arquivos.

## Funcionalidades

- **Processamento automático de dados** de equipamentos e office-boys
- **Cálculo automático de intervalo de datas** baseado nos dados do arquivo
- **Geração de relatório Excel** com duas planilhas:
  - **PRODUTIVIDADE**: Tabela de dados e 4 gráficos de pizza interativos
  - **BASE**: Dados processados completos
- **Gráficos visuais** com cores personalizadas e análise por equipe/tipo de solicitação
- **Interface web simples** para upload de arquivos

## Estrutura do Projeto

```
controladoria-web
├── src
│   ├── app.py                # Aplicação Flask principal
│   ├── controladoria.py      # Lógica de processamento de dados e geração de gráficos
│   ├── templates
│   │   └── index.html        # Interface de upload de arquivos
│   └── static
│       └── style.css         # Estilos da interface web
├── requirements.txt          # Dependências do projeto
└── README.md                 # Documentação do projeto
```

## Requisitos dos Arquivos

A aplicação espera dois arquivos Excel:

1. **Painel de Serviços.xlsx**:
   - Deve conter uma planilha "Sheet1" com cabeçalho na linha 2
   - Colunas necessárias: "Data/Hora Encerramento", "Usuário do Encerramento", "Status", "Tipo de Serviço"

2. **Office-boys.xlsx**:
   - Deve conter colunas "Matrícula" e "Nome"
   - Usado para mapear usuários a nomes e determinar equipes

## Instruções de Instalação

1. **Clone o repositório**:
   ```bash
   git clone <repository-url>
   cd controladoria-web
   ```

2. **Crie um ambiente virtual** (recomendado):
   ```bash
   python -m venv venv
   venv\Scripts\activate  # No Windows
   ```

3. **Instale as dependências**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Execute a aplicação**:
   ```bash
   python src/app.py
   ```
   A aplicação estará disponível em `http://127.0.0.1:5000`

## Como Usar

1. Abra o navegador e acesse `http://127.0.0.1:5000`
2. Selecione o arquivo "Painel de Serviços.xlsx"
3. Selecione o arquivo "Office-boys.xlsx"
4. Clique em "Criar Relatório"
5. O arquivo "RECOLHIMENTO CONTROLADORIA.xlsx" será baixado automaticamente.

## Relatório Gerado

O relatório Excel contém:

### Planilha PRODUTIVIDADE (primeira aba)
- **Tabela de dados** com contagens por equipe, tipo e solicitação
- **4 gráficos de pizza** (10cm x 6cm cada):
  - Recolhimento - 1ª Tentativa (Produtiva vs Improdutiva)
  - Técnica - 1ª Tentativa (Produtiva vs Improdutiva)
  - Recolhimento - Revisita (Produtiva vs Improdutiva)
  - Técnica - Revisita (Produtiva vs Improdutiva)
- **Cores personalizadas**: Rosa forte (produtiva) e Azul escuro (improdutiva)

### Planilha BASE (segunda aba)
- Dados completos processados com colunas adicionais:
  - "Nome do Técnico" (mapeado do arquivo office-boys)
  - "Equipe" (Recolhimento ou Técnica)

## Tecnologias Utilizadas

- **Flask**: Framework web Python
- **Pandas**: Manipulação e análise de dados
- **OpenPyXL**: Leitura/escrita de arquivos Excel e criação de gráficos
- **HTML/CSS**: Interface do usuário

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para submeter pull requests ou abrir issues para sugestões e melhorias.

## Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.