"""Matplotlib grafik modulu - NPK dagilim grafigi."""
import os
os.environ["MPLBACKEND"] = "Agg"

import matplotlib
matplotlib.use("Agg")

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from matplotlib import font_manager
import matplotlib.pyplot as plt


def turkce_font_ayarla():
    """Turkce karakter destekli font ayari."""
    try:
        windows_font = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
        arial_path = os.path.join(windows_font, "arial.ttf")
        if os.path.exists(arial_path):
            font_manager.fontManager.addfont(arial_path)
            plt.rcParams["font.family"] = "Arial"
    except Exception:
        pass


turkce_font_ayarla()


class NPKPastalGrafik(FigureCanvasQTAgg):
    """NPK dagilim pastal grafigi."""

    def __init__(self, n_deger=0, p_deger=0, k_deger=0, ebeveyn=None):
        self.fig = Figure(figsize=(4, 3), dpi=100)
        self.fig.set_facecolor("none")
        super().__init__(self.fig)
        self.setParent(ebeveyn)
        self.grafik_ciz(n_deger, p_deger, k_deger)

    def grafik_ciz(self, n, p, k):
        self.fig.clear()
        ax = self.fig.add_subplot(111)

        toplam = n + p + k
        if toplam == 0:
            ax.text(0.5, 0.5, "Veri Yok", ha="center", va="center",
                    fontsize=12, color="#757575")
            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)
            ax.axis("off")
            self.draw()
            return

        degerler = [n, p, k]
        etiketler = [f"N\n{n} kg", f"P2O5\n{p} kg", f"K2O\n{k} kg"]
        renkler = ["#2E7D32", "#1565C0", "#E65100"]

        wedges, texts, autotexts = ax.pie(
            degerler, labels=etiketler, colors=renkler,
            autopct="%1.0f%%", startangle=90,
            textprops={"fontsize": 9}
        )
        for autotext in autotexts:
            autotext.set_color("white")
            autotext.set_fontweight("bold")

        ax.set_title("NPK Dagilimi", fontsize=11, fontweight="bold", color="#2E7D32")
        self.fig.tight_layout()
        self.draw()

    def guncelle(self, n, p, k):
        self.grafik_ciz(n, p, k)


class NPKCubukGrafik(FigureCanvasQTAgg):
    """NPK ihtiyac vs onerilen cubuk grafigi."""

    def __init__(self, ebeveyn=None):
        self.fig = Figure(figsize=(5, 3), dpi=100)
        self.fig.set_facecolor("none")
        super().__init__(self.fig)
        self.setParent(ebeveyn)
        self._bos_grafik()

    def _bos_grafik(self):
        self.fig.clear()
        ax = self.fig.add_subplot(111)
        ax.text(0.5, 0.5, "Hesaplama Yapilmadi", ha="center", va="center",
                fontsize=11, color="#757575")
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        self.draw()

    def grafik_ciz(self, n_ihtiyac, p_ihtiyac, k_ihtiyac,
                   n_oneri, p_oneri, k_oneri):
        self.fig.clear()
        ax = self.fig.add_subplot(111)

        kategoriler = ["Azot (N)", "Fosfor (P2O5)", "Potasyum (K2O)"]
        ihti_y = [n_ihtiyac, p_ihtiyac, k_ihtiyac]
        oner_y = [n_oneri, p_oneri, k_oneri]

        x = range(len(kategoriler))
        genislik = 0.35

        ax.bar([i - genislik/2 for i in x], ihti_y, genislik,
               label="Ihtiyac (kg/da)", color="#2E7D32", alpha=0.8)
        ax.bar([i + genislik/2 for i in x], oner_y, genislik,
               label="Onerilen (kg/da)", color="#1565C0", alpha=0.8)

        ax.set_xticks(x)
        ax.set_xticklabels(kategoriler, fontsize=9)
        ax.set_ylabel("kg/da", fontsize=9)
        ax.set_title("Ihtiyac vs Onerilen", fontsize=11, fontweight="bold", color="#2E7D32")
        ax.legend(fontsize=8)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        self.fig.tight_layout()
        self.draw()

    def guncelle(self, n_iht, p_iht, k_iht, n_on, p_on, k_on):
        self.grafik_ciz(n_iht, p_iht, k_iht, n_on, p_on, k_on)
