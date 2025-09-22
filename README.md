# Search N Upload

Este projeto automatiza o processo de busca, renomeação e upload de arquivos, utilizando planilhas Excel como base de dados.

## Funcionalidades
- Busca de nomes de diretórios e processos em planilhas Excel
- Busca recursiva de arquivos em diretórios de rede
- Renomeação automática dos arquivos conforme padrão pré estabelecido
- Upload dos arquivos para uma API via requisição HTTP


## Requisitos
- Python 3.10+
- Pacotes: `requests`, `python-dotenv`, `openpyxl`

## Como usar
1. Configure as variáveis de ambiente no arquivo `.env`
2. Ajuste os caminhos das planilhas
3. Execute o script principal:
   ```bash
   python index.py
   ```

## Observações

- O script renomeia os arquivos para o padrão `processo_nome_do_arquivo.pdf`
