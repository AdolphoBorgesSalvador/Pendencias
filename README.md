# Projeto de Análise de Pendências

Este projeto é uma automação para análise e manipulação de dados relacionados a pendências, utilizando Python e bibliotecas como `pandas`. Ele processa arquivos Excel contendo informações de estoque, pedidos e classificações de materiais, gerando tabelas dinâmicas e unindo dados de diversas fontes.

## Estrutura do Projeto

### Diretórios e Arquivos
- **`Pendencias/Arquivos`**: Contém os arquivos Excel necessários para a análise. **(Ignorado no repositório para proteger dados sensíveis).**
- **`pendencias.ipynb`**: Notebook principal com o código para análise e manipulação de dados.

### Principais Arquivos de Dados
- `PEND.XLSX`: Tabela de pendências.
- `zstok.XLSX`: Dados de estoque.
- `fup.XLSX`: Dados de pedidos.
- `Mapa.xlsx`: Consolidado final (gerado pelo script).
- Outros arquivos necessários para as tabelas dinâmicas.

---

## Funcionalidades

1. **Leitura de Arquivos Excel**:
   - Carrega diversas planilhas de dados, como pendências, estoque, e pedidos.
   
2. **Manipulação de Dados**:
   - Geração de tabelas dinâmicas (`pivot_table`) para sumarizar informações.
   - Junções (`merge`) entre diferentes tabelas para consolidar os dados.

3. **Automatização de Processos**:
   - Funções para buscar informações de estoque e similares automaticamente.
   - Geração de arquivo Excel consolidado (`Mapa.xlsx`).

---

## Como Usar

### Requisitos
- Python 3.x
- Bibliotecas necessárias: `pandas`, `openpyxl`

### Configuração
1. Clone este repositório:
   ```bash
   git clone https://github.com/seuusuario/seurepositorio.git
