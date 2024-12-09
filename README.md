# Gerador de Planilhas com IA

Este programa utiliza o Ollama para gerar planilhas Excel automaticamente a partir de descrições em linguagem natural.

## Pré-requisitos

1. Python 3.8 ou superior
2. Ollama instalado e rodando localmente
3. Modelo Mistral (ou outro modelo) instalado no Ollama

## Instalação

1. Clone este repositório
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

3. Certifique-se que o Ollama está rodando localmente na porta padrão (11434)

## Uso

1. Execute o programa:
```bash
python planilha_ia.py
```

2. Digite sua descrição da planilha que deseja gerar
3. Opcionalmente, forneça um nome para o arquivo
4. A planilha será gerada no diretório atual

## Exemplos de Prompts

- "Crie uma planilha de controle financeiro com colunas para data, descrição, valor e categoria"
- "Gere uma lista de tarefas com colunas para título, prioridade, status e prazo"
- "Monte uma planilha de contatos com nome, email, telefone e empresa"

## Observações

- O programa espera que o Ollama retorne dados em formato JSON válido
- As respostas do modelo serão automaticamente convertidas em planilhas Excel
- Os arquivos são salvos com timestamp no nome caso não seja especificado um nome

## Suporte

Para problemas ou sugestões, abra uma issue no repositório. 