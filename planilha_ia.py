import json
import requests
import pandas as pd
from typing import Dict, List, Any
import os
from datetime import datetime

class GeradorPlanilhas:
    def __init__(self, modelo: str = "mistral"):
        """
        Inicializa o gerador de planilhas.
        
        Args:
            modelo (str): Nome do modelo Ollama a ser usado
        """
        self.modelo = modelo
        self.url_base = "http://localhost:11434/api/generate"
        
    def _criar_prompt_template(self, descricao: str) -> str:
        """
        Cria um prompt template específico para gerar dados em formato JSON.
        """
        return f"""Por favor, gere dados para uma planilha com base na seguinte descrição:
{descricao}

Importante: Responda APENAS com um JSON válido seguindo estas regras:
1. O JSON deve ser uma lista de objetos
2. Cada objeto deve ter as mesmas chaves
3. Os valores devem ser apropriados para uma planilha
4. Gere pelo menos 5 itens de exemplo
5. Use apenas strings e números como valores
6. Não inclua comentários ou explicações

Exemplo de formato esperado:
[
  {{"coluna1": "valor1", "coluna2": "valor2"}},
  {{"coluna1": "valor3", "coluna2": "valor4"}}
]

Gere o JSON agora:"""

    def _extrair_json_da_resposta(self, texto: str) -> Dict[str, Any]:
        """
        Tenta extrair um JSON válido da resposta do modelo.
        """
        # Remove possíveis textos antes e depois do JSON
        try:
            # Encontra o primeiro '[' e último ']'
            inicio = texto.find('[')
            fim = texto.rfind(']') + 1
            
            if inicio == -1 or fim == 0:
                raise ValueError("Não foi possível encontrar um JSON válido na resposta")
                
            json_str = texto[inicio:fim]
            return json.loads(json_str)
        except Exception as e:
            raise ValueError(f"Erro ao processar resposta do modelo: {str(e)}\nResposta recebida: {texto}")
            
    def _fazer_requisicao_ollama(self, prompt: str) -> Dict[str, Any]:
        """
        Faz uma requisição para o Ollama e retorna a resposta em formato JSON.
        
        Args:
            prompt (str): Prompt para o modelo
            
        Returns:
            Dict[str, Any]: Resposta do modelo em formato JSON
        """
        headers = {"Content-Type": "application/json"}
        data = {
            "model": self.modelo,
            "prompt": prompt,
            "stream": False,
            "temperature": 0.7,  # Ajustando temperatura para melhor estruturação
            "system": "Você é um assistente especializado em gerar dados estruturados em formato JSON."
        }
        
        try:
            response = requests.post(self.url_base, headers=headers, json=data)
            response.raise_for_status()
            
            # Extrai a resposta do modelo
            resposta_texto = response.json()["response"]
            
            # Tenta extrair e parsear o JSON da resposta
            return self._extrair_json_da_resposta(resposta_texto)
            
        except requests.exceptions.RequestException as e:
            raise ConnectionError(f"Erro ao conectar com o Ollama: {str(e)}")
        except Exception as e:
            raise ValueError(f"Erro ao processar resposta: {str(e)}")
            
    def gerar_planilha(self, prompt: str, nome_arquivo: str = None) -> str:
        """
        Gera uma planilha Excel baseada no prompt fornecido.
        
        Args:
            prompt (str): Descrição do que deve ser gerado na planilha
            nome_arquivo (str, opcional): Nome do arquivo Excel a ser gerado
            
        Returns:
            str: Caminho do arquivo gerado
        """
        # Gera nome do arquivo se não fornecido
        if nome_arquivo is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"planilha_{timestamp}.xlsx"
        elif not nome_arquivo.endswith('.xlsx'):
            nome_arquivo += '.xlsx'
            
        # Cria o prompt estruturado
        prompt_estruturado = self._criar_prompt_template(prompt)
            
        # Obtém dados do modelo
        dados = self._fazer_requisicao_ollama(prompt_estruturado)
        
        if not isinstance(dados, list) or not dados:
            raise ValueError("O modelo não retornou uma lista de dados válida")
            
        # Converte para DataFrame
        df = pd.DataFrame(dados)
        
        # Salva planilha
        df.to_excel(nome_arquivo, index=False, engine='openpyxl')
        
        return nome_arquivo

def main():
    # Exemplo de uso
    gerador = GeradorPlanilhas()
    
    print("🤖 Gerador de Planilhas com IA")
    print("-" * 30)
    
    while True:
        prompt = input("\nDescreva a planilha que você deseja gerar (ou 'sair' para encerrar): ")
        
        if prompt.lower() == 'sair':
            break
            
        try:
            nome_arquivo = input("Nome do arquivo (opcional, pressione Enter para automático): ").strip()
            nome_arquivo = None if not nome_arquivo else nome_arquivo
            
            arquivo_gerado = gerador.gerar_planilha(prompt, nome_arquivo)
            print(f"\n✅ Planilha gerada com sucesso: {arquivo_gerado}")
            
        except Exception as e:
            print(f"\n❌ Erro ao gerar planilha: {str(e)}")
            
    print("\nObrigado por usar o Gerador de Planilhas com IA! 👋")

if __name__ == "__main__":
    main() 