# 📰 Widget News

Widget flutuante de notícias para área de trabalho Windows, com resumo automático via IA (Claude).

## Funcionalidades

- Busca notícias em tempo real via feeds RSS (Brasil, Internacional, Tecnologia)
- Resume manchetes automaticamente usando a API do Claude (Haiku)
- Interface sempre visível, arrastável e sem barra de título do sistema
- Atualização automática a cada 30 minutos
- Atualização manual com botão ⟳

## Tecnologias

- Python 3.11+
- Tkinter (UI)
- feedparser (leitura de RSS)
- Anthropic SDK (resumo com Claude)

## Instalação

```bash
pip install -r requirements.txt
```

## Configuração

Defina a chave da API como variável de ambiente antes de executar:

```powershell
# PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-sua-chave-aqui"
python news_widget.py
```

> A chave **nunca** deve ser commitada no repositório. Use sempre variável de ambiente ou arquivo `.env` listado no `.gitignore`.

## Execução

```bash
python news_widget.py
```

O widget aparece no canto superior direito da tela.

## Feeds monitorados

| Categoria | Portais |
|-----------|---------|
| 🇧🇷 Brasil | G1, Folha de S.Paulo, UOL |
| 🌍 Internacional | BBC, CNN, Reuters |
| 💻 Tecnologia | TechCrunch, The Verge, Wired |

## Personalização

Edite as constantes no topo de `news_widget.py`:

| Constante | Padrão | Descrição |
|-----------|--------|-----------|
| `UPDATE_INTERVAL_MINUTES` | 30 | Intervalo de atualização automática |
| `MAX_ITEMS_PER_FEED` | 5 | Manchetes por portal |
| `WIDGET_WIDTH / HEIGHT` | 390 × 520 | Dimensões do widget |
| `RSS_FEEDS` | — | Dicionário de portais e URLs RSS |

## Estrutura do projeto

```
Widget_News/
├── news_widget.py   # Código principal
├── requirements.txt # Dependências
├── widget.md        # Esta documentação
└── COMO_USAR.md     # Guia detalhado de uso
```
