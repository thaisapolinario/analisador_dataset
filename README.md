# Classificador Automático de Datasets

Este projeto realiza a varredura, classificação e geração automática de relatórios executivos (TXT e XLSX) para datasets armazenados na pasta `data/`.

### 1. Clonar o repositório
```bash
git clone [https://github.com/thaisapolinario/analisador_dataset.git](https://github.com/thaisapolinario/analisador_dataset.git)


cd analisador_dataset
```

### 2. Criar e ativar um ambiente virtual
```bash
python -m venv .venv
# No Windows (PowerShell):
.venv\Scripts\Activate.ps1
# No Linux/Mac:
source .venv/bin/activate
```

### 3. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 4. Executar o script
Coloque os seus datasets dentro da pasta `data/` e rode:
```bash
python classifica_dataset.py
```

Os relatórios `relatorio_resumo_datasets.txt` e `relatorio_grupos_datasets.xlsx` serão gerados automaticamente.