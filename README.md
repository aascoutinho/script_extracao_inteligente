# 📌 Script de Extração e Interpretação Inteligente

**Versão:** 1.0 (Primeira versão)

## 📋 Sobre o Projeto

Este programa foi desenvolvido por **Antonio Augusto Silva**, área Comercial do **Grupo DR** ([www.grupodr.com.br](http://www.grupodr.com.br)), com o objetivo de realizar automaticamente a extração e interpretação estruturada dos dados encontrados na coluna **"HISTORICO"** de planilhas Excel padrão (`.xlsx`).

O script é capaz de identificar automaticamente 16 padrões distintos pré-definidos, organizando as informações claramente em novas colunas no mesmo arquivo Excel.

## 🚀 Tecnologias

* **Python 3.x**
* **openpyxl** (para manipulação de arquivos Excel mantendo formatação original)
* **tqdm** (para barra de progresso no prompt)

## ✅ Como Executar

### Passo 1 - Instalação do Python

Baixe e instale o Python (marcando a opção "Add Python to PATH"):

* [Baixar Python](https://www.python.org/downloads/)

### Passo 2 - Instalação das dependências

Abra o Prompt de Comando (`cmd`) e execute:

```bash
pip install openpyxl tqdm
```

### Passo 3 - Executar o script

* Copie os arquivos `script_extracao_inteligente.py` e `executar_extracao.bat` para a pasta desejada.
* Coloque as planilhas Excel (ex.: `01 - Janeiro.xlsx`) na mesma pasta.
* **Importante**: Para processar diferentes meses, altere o nome do arquivo Excel diretamente no script Python (`script_extracao_inteligente.py`). Por exemplo:

```python
arquivo = '02 - Fevereiro.xlsx'  # Altere aqui o mês desejado
```

* Dê dois cliques no arquivo:

```
executar_extracao.bat
```

O resultado será salvo automaticamente em uma nova planilha:

```
02 - Fevereiro_Extraido_Inteligente.xlsx
```

## 📑 Estrutura do resultado

Após a execução, serão geradas novas colunas na planilha:

| Coluna         | Descrição                                      |
| -------------- | ---------------------------------------------- |
| Padrão         | Tipo de padrão identificado automaticamente.   |
| Código         | Código referente à operação (quando aplicável) |
| Empresa/Pessoa | Nome da empresa ou pessoa envolvida.           |
| Descrição      | Breve descrição da operação.                   |
| Pedido         | Número do pedido (quando aplicável).           |
| Observação     | Informações adicionais encontradas.            |

## 📂 Estrutura Recomendada de Pastas

```text
/script_extracao_inteligente/
├── scripts/
│   └── script_extracao_inteligente.py
├── planilhas/
│   └── 01 - Janeiro.xlsx
├── executar_extracao.bat
├── README.md
├── .gitignore
└── LICENSE (opcional)
```

## 🛡️ .gitignore sugerido

```gitignore
__pycache__/
*.pyc
*.pyo
*.pyd
*.log
*.xlsx
.env
.DS_Store
```

## 🏁 Criando uma versão (tag) no GitHub

```bash
git tag v1.0
git push origin v1.0
```

## 📞 Suporte

Para suporte ou mais informações:

**Antonio Augusto Silva**
**Área Comercial - Grupo DR**
📞 **+55 11 4712-2231**
🌐 [www.grupodr.com.br](https://www.grupodr.com.br)
