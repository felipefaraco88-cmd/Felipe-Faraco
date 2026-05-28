# 📰 News Widget — Guia de Configuração

## O que é
Um widget flutuante que fica sempre visível na sua área de trabalho,
buscando notícias de portais brasileiros, internacionais e de tecnologia,
e resumindo tudo com a IA do Claude.

---

## 1. Pré-requisitos

- **Python 3.11+** instalado → https://www.python.org/downloads/
- **VS Code** com a extensão **Python** da Microsoft
- Sua **chave da API da Anthropic** → https://console.anthropic.com/

---

## 2. Configuração no VS Code

### 2.1 Abra a pasta do projeto
```
File → Open Folder → selecione a pasta onde salvou os arquivos
```

### 2.2 Abra o terminal integrado
```
Terminal → New Terminal   (ou Ctrl + `)
```

### 2.3 Instale as dependências
```bash
pip install -r requirements.txt
```

### 2.4 Configure sua chave da API

**Opção A — Variável de ambiente (recomendado):**
No terminal do VS Code:
```bash
# Windows (PowerShell)
$env:ANTHROPIC_API_KEY = "sk-ant-sua-chave-aqui"
python news_widget.py

# Windows (cmd)
set ANTHROPIC_API_KEY=sk-ant-sua-chave-aqui
python news_widget.py
```

**Opção B — Diretamente no código:**
Abra `news_widget.py` e edite a linha:
```python
CLAUDE_API_KEY = "sk-ant-sua-chave-aqui"
```

---

## 3. Executar o widget

```bash
python news_widget.py
```

O widget aparecerá no canto superior direito da tela.

---

## 4. Como usar o widget

| Ação | Como fazer |
|------|-----------|
| Mover | Clique e arraste pela barra do título |
| Atualizar manualmente | Clique no botão **⟳** |
| Fechar | Clique no botão **✕** |
| Trocar categoria | Clique nas abas (Brasil / Internacional / Tecnologia) |

A atualização automática ocorre a cada **30 minutos**.

---

## 5. Iniciar automaticamente com o Windows (opcional)

1. Pressione `Win + R`, digite `shell:startup` e pressione Enter
2. Crie um arquivo `news_widget.bat` nessa pasta com o conteúdo:

```bat
@echo off
set ANTHROPIC_API_KEY=sk-ant-sua-chave-aqui
cd /d "C:\caminho\para\sua\pasta"
python news_widget.py
```

O widget vai iniciar automaticamente toda vez que você ligar o computador.

---

## 6. Personalização

No arquivo `news_widget.py`, você pode ajustar:

| Constante | O que faz | Padrão |
|-----------|-----------|--------|
| `UPDATE_INTERVAL_MINUTES` | Intervalo de atualização | 30 min |
| `MAX_ITEMS_PER_FEED` | Manchetes por portal | 5 |
| `WIDGET_WIDTH / HEIGHT` | Tamanho do widget | 390 × 520 px |
| `RSS_FEEDS` | Portais monitorados | Ver seção abaixo |

### Adicionar/remover portais
Edite o dicionário `RSS_FEEDS` no início do script. Qualquer portal com RSS funciona.
Exemplos de feeds extras:
- ESPN Brasil: `https://www.espn.com.br/rss/`
- Globo Esporte: `https://globoesporte.globo.com/rss/`
- MIT Tech Review: `https://www.technologyreview.com/feed/`
