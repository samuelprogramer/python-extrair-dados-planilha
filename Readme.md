# Processador de Planilha

Este script em Python processa uma planilha Excel específica, extrai informações relevantes e as salva em um arquivo JSON.

## Requisitos

- Python 3.x
- Bibliotecas necessárias (`openpyxl`)

Você pode instalar as bibliotecas usando o seguinte comando:

```bash
pip install openpyxl
```

## Uso
Coloque sua planilha no mesmo diretório do script com o nome 'planilhaQuestoes.xlsx'.

Execute o script Python:
```bash
python processador_planilha.py
```

O script processará a planilha e gerará um arquivo JSON chamado output.json no mesmo diretório.

### Detalhes do Código

O código consiste em várias funções para manipulação da planilha, incluindo:

 - `isResponse`: Verifica se uma célula é uma resposta correta.
 - `concatena_valores_celulas`: Concatena os valores de duas células.
 - `get_cell_value`: Obtém o valor de uma célula.
 - `process_line`: Processa uma linha da planilha.
