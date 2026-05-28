"""
📰 News Widget — Bot de Notícias para Área de Trabalho
Busca notícias via RSS e resume com a API do Claude.

Configuração:
  1. Defina a variável de ambiente ANTHROPIC_API_KEY com sua chave
     ou coloque-a direto na constante CLAUDE_API_KEY abaixo.
  2. pip install feedparser anthropic
  3. python news_widget.py
"""

import tkinter as tk
from tkinter import ttk
import feedparser
import anthropic
import threading
import os
from datetime import datetime

# ─── Configuração ────────────────────────────────────────────────────────────
CLAUDE_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "coloque-sua-chave-aqui")
UPDATE_INTERVAL_MINUTES = 30          # Intervalo de atualização automática
MAX_ITEMS_PER_FEED = 5               # Manchetes por feed
WIDGET_WIDTH  = 390
WIDGET_HEIGHT = 520

# ─── Feeds RSS ───────────────────────────────────────────────────────────────
RSS_FEEDS = {
    "🇧🇷 Brasil": [
        ("G1",    "https://g1.globo.com/rss/g1/"),
        ("Folha", "https://feeds.folha.uol.com.br/emcimadahora/rss091.xml"),
        ("UOL",   "https://rss.uol.com.br/feed/noticias.xml"),
    ],
    "🌍 Internacional": [
        ("BBC",     "https://feeds.bbci.co.uk/news/rss.xml"),
        ("CNN",     "http://rss.cnn.com/rss/edition.rss"),
        ("Reuters", "https://feeds.reuters.com/reuters/topNews"),
    ],
    "💻 Tecnologia": [
        ("TechCrunch", "https://techcrunch.com/feed/"),
        ("The Verge",  "https://www.theverge.com/rss/index.xml"),
        ("Wired",      "https://www.wired.com/feed/rss"),
    ],
}

# ─── Cores ───────────────────────────────────────────────────────────────────
BG_DARK    = "#0d1117"
BG_TITLE   = "#161b22"
BG_TAB     = "#21262d"
FG_WHITE   = "#e6edf3"
FG_GRAY    = "#8b949e"
FG_ACCENT  = "#58a6ff"
FG_RED     = "#ff7b72"
FG_GREEN   = "#3fb950"


# ─── Funções de negócio ──────────────────────────────────────────────────────

def fetch_news(feed_url: str, max_items: int = MAX_ITEMS_PER_FEED) -> list[str]:
    """Retorna lista de manchetes + trecho do feed RSS."""
    try:
        feed = feedparser.parse(feed_url)
        items = []
        for entry in feed.entries[:max_items]:
            title   = entry.get("title", "").strip()
            summary = entry.get("summary", entry.get("description", "")).strip()
            # Remove HTML tags simples
            import re
            summary = re.sub(r"<[^>]+>", "", summary)[:180]
            items.append(f"• {title}. {summary}")
        return items if items else ["(sem itens no feed)"]
    except Exception as exc:
        return [f"Erro ao buscar feed: {exc}"]


def summarize_with_claude(news_text: str, category: str, api_key: str) -> str:
    """Envia manchetes ao Claude e retorna resumo em bullet points."""
    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=600,
            messages=[
                {
                    "role": "user",
                    "content": (
                        f"Aqui estão as principais manchetes de {category}.\n"
                        "Faça um resumo CONCISO em português dos 4-5 pontos mais importantes.\n"
                        "Use bullet points curtos (máx. 2 linhas cada).\n"
                        "Seja direto — sem introdução nem conclusão.\n\n"
                        f"Manchetes:\n{news_text}"
                    ),
                }
            ],
        )
        return msg.content[0].text.strip()
    except anthropic.AuthenticationError:
        return "❌ Chave da API inválida.\nVerifique ANTHROPIC_API_KEY."
    except Exception as exc:
        return f"❌ Erro na API do Claude:\n{exc}"


# ─── Widget ──────────────────────────────────────────────────────────────────

class NewsWidget:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Notícias")
        self.root.geometry(f"{WIDGET_WIDTH}x{WIDGET_HEIGHT}")
        self.root.configure(bg=BG_DARK)
        self.root.attributes("-topmost", True)
        self.root.overrideredirect(True)   # Remove barra de título do SO

        self._drag_x = 0
        self._drag_y = 0
        self._loading = False

        self._build_ui()
        self._position_window()
        self._schedule_refresh(delay_ms=100)   # Primeira carga imediata

    # ── Construção da UI ──────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_titlebar()
        self._build_status_bar()
        self._build_notebook()

    def _build_titlebar(self):
        bar = tk.Frame(self.root, bg=BG_TITLE, height=38)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        tk.Label(
            bar, text="📰  Notícias do Dia",
            bg=BG_TITLE, fg=FG_WHITE,
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left", padx=12, pady=8)

        # Botão fechar
        tk.Button(
            bar, text="✕", bg=BG_TITLE, fg=FG_RED,
            font=("Segoe UI", 10, "bold"), bd=0,
            activebackground=BG_TITLE, activeforeground="#ff4444",
            cursor="hand2", command=self.root.destroy,
        ).pack(side="right", padx=10)

        # Botão atualizar
        self.refresh_btn = tk.Button(
            bar, text="⟳", bg=BG_TITLE, fg=FG_ACCENT,
            font=("Segoe UI", 13), bd=0,
            activebackground=BG_TITLE, activeforeground=FG_WHITE,
            cursor="hand2", command=self._manual_refresh,
        )
        self.refresh_btn.pack(side="right", padx=4)

        # Arraste pelo título
        for w in (bar,):
            w.bind("<Button-1>",   self._drag_start)
            w.bind("<B1-Motion>",  self._drag_move)

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="Carregando…")
        tk.Label(
            self.root, textvariable=self.status_var,
            bg=BG_DARK, fg=FG_GRAY,
            font=("Segoe UI", 8),
        ).pack(pady=(2, 0))

    def _build_notebook(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook",        background=BG_DARK,  borderwidth=0)
        style.configure("TNotebook.Tab",    background=BG_TAB,   foreground=FG_GRAY,
                         padding=[10, 5],   font=("Segoe UI", 9))
        style.map("TNotebook.Tab",
                  background=[("selected", BG_TITLE)],
                  foreground=[("selected", FG_WHITE)])

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=6, pady=6)

        self.text_areas: dict[str, tk.Text] = {}

        for category in RSS_FEEDS:
            frame = tk.Frame(nb, bg=BG_DARK)
            nb.add(frame, text=category)

            text = tk.Text(
                frame,
                bg=BG_DARK, fg=FG_WHITE,
                font=("Segoe UI", 9),
                wrap="word", bd=0,
                padx=10, pady=8,
                cursor="arrow",
                state="disabled",
                spacing1=2, spacing3=4,
            )
            scrollbar = tk.Scrollbar(frame, orient="vertical",
                                     command=text.yview, bg=BG_TAB,
                                     troughcolor=BG_DARK, bd=0)
            text.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            text.pack(fill="both", expand=True)

            # Tags de formatação
            text.tag_configure("accent",  foreground=FG_ACCENT)
            text.tag_configure("gray",    foreground=FG_GRAY)
            text.tag_configure("green",   foreground=FG_GREEN)

            self.text_areas[category] = text

    # ── Posicionamento e arraste ───────────────────────────────────────────────

    def _position_window(self):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        x  = sw - WIDGET_WIDTH - 20
        y  = 50
        self.root.geometry(f"{WIDGET_WIDTH}x{WIDGET_HEIGHT}+{x}+{y}")

    def _drag_start(self, event):
        self._drag_x = event.x
        self._drag_y = event.y

    def _drag_move(self, event):
        x = self.root.winfo_pointerx() - self._drag_x
        y = self.root.winfo_pointery() - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    # ── Lógica de atualização ─────────────────────────────────────────────────

    def _set_text(self, category: str, content: str, tags: list = None):
        """Atualiza o conteúdo de uma aba de forma thread-safe."""
        def _update():
            area = self.text_areas[category]
            area.configure(state="normal")
            area.delete("1.0", "end")
            area.insert("end", content)
            area.configure(state="disabled")
        self.root.after(0, _update)

    def _fetch_all(self):
        """Roda em thread separada: busca RSS e resume com Claude."""
        self._loading = True
        self.root.after(0, lambda: self.refresh_btn.configure(text="…"))

        for category, feeds in RSS_FEEDS.items():
            self._set_text(category, "⏳ Buscando notícias…")

            lines = []
            for source, url in feeds:
                items = fetch_news(url)
                lines.append(f"\n[{source}]")
                lines.extend(items)

            raw_text = "\n".join(lines)
            summary  = summarize_with_claude(raw_text, category, CLAUDE_API_KEY)

            hora = datetime.now().strftime("%H:%M")
            self._set_text(category, f"🕐 Atualizado às {hora}\n\n{summary}")

        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.root.after(0, lambda: self.status_var.set(f"Última atualização: {now}"))
        self.root.after(0, lambda: self.refresh_btn.configure(text="⟳"))
        self._loading = False

    def _manual_refresh(self):
        if not self._loading:
            threading.Thread(target=self._fetch_all, daemon=True).start()

    def _schedule_refresh(self, delay_ms: int = None):
        ms = delay_ms if delay_ms is not None else UPDATE_INTERVAL_MINUTES * 60 * 1000
        self.root.after(ms, self._auto_refresh)

    def _auto_refresh(self):
        self._manual_refresh()
        self._schedule_refresh()   # Reagenda

    # ── Loop principal ────────────────────────────────────────────────────────

    def run(self):
        self.root.mainloop()


# ─── Entrypoint ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    widget = NewsWidget()
    widget.run()
