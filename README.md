# Script de Processamento de Arquivos Excel

Este script Python permite processar arquivos Excel, realizar consultas ao banco de dados SQL Server e mover as linhas correspondentes para as pastas "Processado" ou "Rejeitado", de acordo com determinadas condições.

## Requisitos

- Python 3.x
- Bibliotecas Python: pyodbc, pandas, openpyxl

Certifique-se de ter as bibliotecas necessárias instaladas antes de executar o script.

## Configuração

1. Defina as configurações do banco de dados SQL Server no código, fornecendo o nome do servidor, nome do banco de dados e credenciais (se aplicável).

2. Defina os caminhos das pastas onde os arquivos Excel serão lidos, processados, rejeitados e onde os logs serão armazenados.

3. Crie as tabelas necessárias no banco de dados usando o script SQL fornecido no arquivo `create_tables.sql`. Certifique-se de ter o SQL Server configurado corretamente e tenha privilégios para criar tabelas e inserir dados.

## Utilização

1. Coloque os arquivos Excel na pasta especificada (`caminho_pasta`).

2. Execute o script Python `main.py`.

3. O script lerá cada arquivo Excel, realizará consultas ao banco de dados, atualizará as linhas correspondentes e moverá os arquivos para as pastas "Processado" ou "Rejeitado".

4. Os logs de processamento serão registrados nos arquivos `log_basico.txt` e `log_<arquivo_excel>.txt`, fornecendo informações sobre as linhas processadas, atualizadas e rejeitadas.

Certifique-se de ter as permissões adequadas nos diretórios e arquivos necessários.

## Observações

- O script assume que os arquivos Excel têm as colunas "Sequencial" e "Cnpj" que serão usadas para a consulta e atualização no banco de dados.

- Certifique-se de que as tabelas e colunas do banco de dados correspondam corretamente às consultas e atualizações realizadas no script.

- Verifique e ajuste o script de acordo com as necessidades específicas do seu ambiente e dados.

