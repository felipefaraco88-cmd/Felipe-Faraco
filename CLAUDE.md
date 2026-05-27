# Assistente Fiscal NF-e / NFS-e — TECNOGLOBO EQUIPAMENTOS LTDA

## Visão Geral

Assistente para montagem de rascunho de Nota Fiscal Eletrônica (NF-e / NFS-e). A interface principal é um arquivo HTML standalone (`NF_ASSIST.html`) que roda diretamente no navegador sem dependências externas. O Python (`NF_ASSIST_CLD.py`) serve como lançador e contém o gerador DANFE em PDF via `fpdf2`.

## Arquivos

| Arquivo | Descrição |
|---|---|
| `NF_ASSIST.html` | Interface principal — abrir no browser |
| `NF_ASSIST_CLD.py` | Lançador Python + gerador DANFE PDF (fallback Tkinter) |

## Como Usar

### Opção 1 — Direto no browser (recomendado)
Abra `NF_ASSIST.html` no navegador (duplo clique ou `Ctrl+O`).

### Opção 2 — Via Python
```bash
python NF_ASSIST_CLD.py
```
O script detecta o `NF_ASSIST.html` e o abre automaticamente no navegador.

### Dependências Python (só para fallback Tkinter / Excel)
```bash
pip install openpyxl fpdf2 requests
```

## Fluxo do Assistente (7 Blocos)

1. **Tipo de Nota** — NF-e venda, NFS-e serviço, devolução, remessa, etc. (inclui campo "Outros")
2. **Destinatário** — Razão social, CPF/CNPJ, endereço, e-mail
3. **Produtos / Equipamentos** — Múltiplos itens com descrição, NCM, quantidade, valor unitário e número de série
4. **Códigos Fiscais** — NCM, CFOP (sugestão automática por NCM), natureza da operação
5. **Condições** — Forma de pagamento, modalidade de frete, transportadora
6. **Observações** — Texto livre para o campo de informações adicionais
7. **Resumo Final** — Revisão completa + exportação

## Exportações Disponíveis

| Formato | Como funciona |
|---|---|
| **PDF (DANFE)** | Abre nova aba com DANFE formatado e aciona `window.print()` → salvar como PDF |
| **XML (NF-e)** | Gera XML padrão NF-e v4.00 (estrutura SEFAZ/ABNT) e faz download do `.xml` |
| **HTML** | Salva o DANFE como arquivo `.html` local |
| **Excel** | Exportação via Python (`openpyxl`) — disponível no fallback Tkinter |

## Dados do Emitente

Configurados diretamente no topo de `NF_ASSIST_CLD.py` no dicionário `EMITENTE` e no objeto `E` dentro de `NF_ASSIST.html`:

```python
# NF_ASSIST_CLD.py
EMITENTE = {
    "razao_social": "TECNOGLOBO EQUIPAMENTOS LTDA",
    "cnpj":         "07.700.041/0001-84",
    "uf":           "PR",
    ...
}
```

Para alterar os dados da empresa, edite **ambos** os arquivos.

## Paleta de Cores

| Variável | Hex | Uso |
|---|---|---|
| `--primary` | `#E65100` | Laranja escuro — títulos de seção |
| `--secondary` | `#5D4037` | Marrom escuro — cabeçalho, labels |
| `--accent` | `#FFB74D` | Laranja claro — destaques |
| `--success` | `#388E3C` | Verde — totais, botões de exportação |
| `--error` | `#D32F2F` | Vermelho — avisos críticos |

## Observações Importantes

- **Documento sem valor fiscal** — este assistente gera apenas rascunhos para conferência interna.
- O XML exportado usa `tpAmb=2` (homologação). Para produção, alterar para `tpAmb=1` e preencher todos os campos obrigatórios da SEFAZ.
- Confirme sempre NCM e CFOP com o contador antes de emitir a nota oficial.
- O campo NCM no Bloco 4 sugere CFOPs automaticamente para os principais equipamentos de informática (computadores, impressoras, roteadores, monitores, notebooks).
