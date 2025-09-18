# ───── 模組匯入 ─────
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageFilter
import pygame
import yfinance as yf
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from datetime import datetime

# ───── 初始化 Pygame Mixer ─────
pygame.mixer.init()

# ───── Matplotlib 中文與負號設定 ─────
matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
matplotlib.rcParams["axes.unicode_minus"] = False

# ───── 全域參數 ─────
risk_free_rate = 0
is_music_playing = False

# ───── 共用工具函式 ─────
def format_ticker(raw: str, market: str) -> str:
    raw = raw.strip().upper()
    if market == "美國":
        return raw
    return raw + {"台灣": ".TW", "日本": ".T", "英國": ".L"}.get(market, "")

def flatten(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df

# ───── 抓取股票資料並計算績效指標 ─────
def fetch_price_and_metrics(ticker: str, start_date=None, end_date=None):
    raw_df = yf.download(ticker, period="max", progress=False, auto_adjust=False)
    df = flatten(raw_df)
    if df.empty or "Adj Close" not in df.columns:
        raise ValueError("抓不到資料或缺少 Adjusted Close 欄位")
    df = df.reset_index()

    if start_date:
        try:
            sd = pd.to_datetime(start_date)
        except:
            raise ValueError("開始日期格式錯誤，請使用 YYYY-MM-DD")
        df = df[df["Date"] >= sd]
    if end_date:
        try:
            ed = pd.to_datetime(end_date)
        except:
            raise ValueError("結束日期格式錯誤，請使用 YYYY-MM-DD")
        df = df[df["Date"] <= ed]

    if df.empty:
        raise ValueError("指定期間內無成交資料")

    df["Return"] = df["Adj Close"].pct_change()
    df.dropna(subset=["Return"], inplace=True)
    df["CumReturn"] = (1 + df["Return"]).cumprod() - 1

    t0, t1 = df["Date"].iloc[0], df["Date"].iloc[-1]
    years = (t1 - t0).days / 365.25
    cagr = (df["Adj Close"].iloc[-1] / df["Adj Close"].iloc[0]) ** (1/years) - 1
    std_annual = df["Return"].std() * np.sqrt(252)
    sharpe = (cagr - risk_free_rate) / std_annual

    df["Downside"] = df["Return"].clip(upper=0)
    d_std = df["Downside"].std() * np.sqrt(252)
    sortino = (cagr - risk_free_rate) / d_std if d_std else np.nan

    metrics = {
        "start": t0.date(),
        "end": t1.date(),
        "years": years,
        "cum": df["CumReturn"].iloc[-1] * 100,
        "ann": cagr * 100,
        "std": std_annual * 100,
        "sharpe": sharpe,
        "sortino": sortino
    }
    return df, metrics

# ───── 「查詢股票資料」功能 ─────
def show_stock():
    raw = entry_stock.get().strip()
    if not raw:
        messagebox.showerror("錯誤", "請輸入股票代號")
        return
    ticker = format_ticker(raw, combo_market_stock.get())

    start_date = entry_start.get().strip()
    end_date = entry_end.get().strip()

    try:
        df, m = fetch_price_and_metrics(
            ticker,
            start_date if start_date else None,
            end_date if end_date else None
        )
    except Exception as e:
        messagebox.showerror("下載失敗", str(e))
        return

    ax_stock.clear()
    ax_stock.plot(df["Date"], df["CumReturn"] * 100, color="#1f77b4")
    ax_stock.set_title(f"{ticker} 累積報酬率 (%)")
    ax_stock.set_xlabel("日期")
    ax_stock.set_ylabel("累積報酬率 (%)")
    ax_stock.grid(True)
    fig_stock.autofmt_xdate()
    canvas_stock.draw()

    txt = (
        f"股票：{ticker}\n"
        f"查詢期間：{m['start']} → {m['end']} ({m['years']:.2f} 年)\n"
        f"累積報酬率：{m['cum']:.2f}%\n"
        f"年化報酬率：{m['ann']:.2f}%\n"
        f"年化標準差：{m['std']:.2f}%\n"
        f"Sharpe 比率：{m['sharpe']:.2f}\n"
        f"Sortino 比率：{m['sortino']:.2f}"
    )
    txt_metrics_stock.config(state="normal")
    txt_metrics_stock.delete("1.0", tk.END)
    txt_metrics_stock.insert(tk.END, txt)
    txt_metrics_stock.config(state="disabled")

# ───── 「退休試算」功能 ─────
def calculate_retirement():
    try:
        expense = float(entry_expense.get())
        infl = float(entry_infl.get()) / 100
    except ValueError:
        messagebox.showerror("錯誤", "年支出或通膨率必須為數字")
        return

    raw = entry_tic_ret.get().strip()
    if not raw:
        messagebox.showerror("錯誤", "請輸入股票代號")
        return
    ticker = format_ticker(raw, combo_market_ret.get())

    start_date = entry_ret_start.get().strip()
    end_date = entry_ret_end.get().strip()

    try:
        df, m = fetch_price_and_metrics(
            ticker,
            start_date if start_date else None,
            end_date if end_date else None
        )
    except Exception as e:
        messagebox.showerror("下載失敗", str(e))
        return

    real_withdraw = (m["ann"] / 100) - infl
    if real_withdraw <= 0:
        messagebox.showwarning("無法計算", "年化報酬不足以支付通膨率")
        return
    need_capital = expense / real_withdraw

    ax_ret.clear()
    ax_ret.plot(df["Date"], df["Adj Close"], color="#ff7f0e")
    ax_ret.set_title(f"{ticker} 調整後收盤價")
    ax_ret.set_xlabel("日期")
    ax_ret.set_ylabel("價格")
    ax_ret.grid(True)
    fig_ret.autofmt_xdate()
    canvas_ret.draw()

    txt = (
        f"股票：{ticker}\n"
        f"查詢期間：{m['start']} → {m['end']} ({m['years']:.2f} 年)\n"
        f"年化報酬率：{m['ann']:.2f}%\n"
        f"預估通膨：{infl * 100:.2f}%\n"
        f"安全提領率：{real_withdraw * 100:.2f}%\n"
        f"年支出：{expense:,.0f} 元\n"
        f"所需退休資產：約 {need_capital:,.0f} 元"
    )
    txt_result_ret.config(state="normal")
    txt_result_ret.delete("1.0", tk.END)
    txt_result_ret.insert(tk.END, txt)
    txt_result_ret.config(state="disabled")

# ───── 「股票查詢頁面」說明按鈕函式 ─────
def show_stock_cum_info():
    message = (
        "累積報酬率：投資期間的總報酬百分比。\n"
        "例如：從 100 → 150，即為 +50%。"
    )
    messagebox.showinfo("什麼是累積報酬率？", message)

def show_stock_cagr_info():
    message = (
        "年化報酬率 (CAGR)：表示平均每年的複利報酬率。\n"
        "理解為每年固定成長 x%，才能在期末得到這段期間的總報酬。"
    )
    messagebox.showinfo("什麼是年化報酬率？", message)

def show_stock_std_info():
    message = (
        "年化標準差：衡量每年報酬的波動度，是常見的風險指標。\n"
        "數值越高，代表報酬波動越大，風險越高。"
    )
    messagebox.showinfo("什麼是年化標準差？", message)

def show_stock_sharpe_info():
    message = (
        "Sharpe 比率：衡量風險調整後的報酬。\n"
        "計算方式 = (年化報酬率 - 無風險利率) ÷ 年化標準差。\n"
        "數值越高表示在相同風險下，獲得的報酬越多。"
    )
    messagebox.showinfo("什麼是 Sharpe 比率？", message)

def show_stock_sortino_info():
    message = (
        "Sortino 比率：只考慮「下跌」風險的報酬風險比。\n"
        "相較於 Sharpe，更注重負向波動帶來的風險影響。"
    )
    messagebox.showinfo("什麼是 Sortino 比率？", message)

# ───── 「畫面切換」與「返回首頁」函式 ─────
def go_home():
    for f in (stock_query_frame, retirement_frame):
        f.pack_forget()
    main_canvas.pack(fill="both", expand=True)

def open_stock_query():
    main_canvas.pack_forget()
    retirement_frame.pack_forget()
    stock_query_frame.pack(fill="both", expand=True)

def open_retirement():
    main_canvas.pack_forget()
    stock_query_frame.pack_forget()
    retirement_frame.pack(fill="both", expand=True)

# ───── 建立主視窗 ─────
root = tb.Window(themename="cosmo")
root.title("投資助手")
root.geometry("1280x800")

# ───── 設定全域字型 & 主題 ─────
style = tb.Style()
default_font = tkFont.nametofont("TkDefaultFont")
default_font.configure(family="Noto Sans TC", size=11)

# ───── 首頁畫布與背景 ─────
main_canvas = tk.Canvas(root, highlightthickness=0)
main_canvas.pack(fill="both", expand=True)

original_bg = Image.open("background4.jpg").filter(ImageFilter.GaussianBlur(radius=2))
bg_pre = original_bg.resize((1280, 800))
bg_photo = ImageTk.PhotoImage(bg_pre)
main_canvas.bg_photo = bg_photo
bg_image_id = main_canvas.create_image(0, 0, image=bg_photo, anchor="nw")

initial_title_font = ("Noto Sans TC", 36, "bold")
title_shadow = main_canvas.create_text(
    0, 0, text="投資助手", font=initial_title_font, fill="#484848", anchor="center"
)
title_id = main_canvas.create_text(
    0, 0, text="投資助手", font=initial_title_font, fill="#ffffff", anchor="center"
)

btn_frame = tb.Frame(main_canvas, style="TFrame")
btn1 = tb.Button(
    btn_frame,
    text="查詢股票資料",
    bootstyle="secondary",
    width=20,
    cursor="hand2",
    takefocus=False,
    command=open_stock_query
)
btn2 = tb.Button(
    btn_frame,
    text="存多少錢可能退休？",
    bootstyle="secondary",
    width=20,
    cursor="hand2",
    takefocus=False,
    command=open_retirement
)
btn1.pack(pady=(0, 8))
btn2.pack()
btn_window_id = main_canvas.create_window(0, 0, window=btn_frame, anchor="n")

speaker_canvas = tk.Canvas(main_canvas, width=50, height=50, bd=0, highlightthickness=0)
speaker_on_icon = ImageTk.PhotoImage(
    Image.open("speakeron1-removebg-preview.png").resize((40, 40))
)
speaker_off_icon = ImageTk.PhotoImage(
    Image.open("speakeroff-removebg-preview.png").resize((40, 40))
)
speaker_icon = speaker_canvas.create_image(25, 25, image=speaker_off_icon)

def toggle_music(event=None):
    global is_music_playing
    if is_music_playing:
        pygame.mixer.music.stop()
        is_music_playing = False
        speaker_canvas.itemconfig(speaker_icon, image=speaker_off_icon)
    else:
        try:
            pygame.mixer.music.load("Moon.mp3")
            pygame.mixer.music.set_volume(0.5)
            pygame.mixer.music.play(-1)
            is_music_playing = True
            speaker_canvas.itemconfig(speaker_icon, image=speaker_on_icon)
        except Exception as e:
            print("Pygame load error:", e)
            messagebox.showerror("音樂檔案錯誤", f"載入失敗：{e}")

speaker_canvas.bind("<Button-1>", toggle_music)
speaker_id = main_canvas.create_window(0, 0, window=speaker_canvas, anchor="ne")

def update_layout(event=None):
    w = main_canvas.winfo_width()
    h = main_canvas.winfo_height()
    if w < 100 or h < 100:
        return

    try:
        resized = original_bg.resize((w, h))
        new_bg = ImageTk.PhotoImage(resized)
        main_canvas.itemconfig(bg_image_id, image=new_bg)
        main_canvas.bg_photo = new_bg

        title_size = max(24, min(48, int(h * 0.06)))
        dyn_title_font = ("Noto Sans TC", title_size, "bold")
        main_canvas.itemconfig(title_shadow, font=dyn_title_font)
        main_canvas.itemconfig(title_id, font=dyn_title_font)

        btn_font_size = max(10, int(h * 0.025))
        dyn_btn_font = ("Noto Sans TC", btn_font_size)
        btn_pad_y = int(h * 0.015)
        style.configure("Dyn.TButton", font=dyn_btn_font, padding=(20, btn_pad_y))
        btn1.configure(style="Dyn.TButton")
        btn2.configure(style="Dyn.TButton")

        main_canvas.coords(title_shadow, w // 2 + 3, int(h * 0.12))
        main_canvas.coords(title_id, w // 2, int(h * 0.11))
        main_canvas.coords(btn_window_id, w // 2, int(h * 0.25))
        main_canvas.coords(speaker_id, w - 30, 30)

    except Exception as e:
        print("resize error:", e)

main_canvas.bind("<Configure>", update_layout)

# ───── 「股票查詢畫面」UI ─────
stock_query_frame = ttk.Frame(root, padding=20)
pw_stock = ttk.PanedWindow(stock_query_frame, orient="horizontal")
pw_stock.pack(fill="both", expand=True)

# 左側：輸入欄位
left_stock = ttk.Frame(pw_stock, padding=10)
pw_stock.add(left_stock, weight=1)

# 「股票查詢頁面」左上：返回首頁按鈕
ttk.Button(left_stock, text="← 返回首頁", command=go_home).pack(anchor="w", pady=(0, 10))

inputs_stock = ttk.LabelFrame(left_stock, text="輸入參數", padding=10)
inputs_stock.pack(fill="x", pady=(0, 10))

ttk.Label(inputs_stock, text="股票代號：").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
entry_stock = ttk.Entry(inputs_stock, width=20)
entry_stock.grid(row=0, column=1, sticky="we", pady=4)

ttk.Label(inputs_stock, text="市場：").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
combo_market_stock = ttk.Combobox(
    inputs_stock, values=["美國", "台灣", "日本", "英國"], state="readonly", width=18
)
combo_market_stock.current(0)
combo_market_stock.grid(row=1, column=1, sticky="we", pady=4)

ttk.Label(inputs_stock, text="開始日期 (YYYY-MM-DD)：").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
entry_start = ttk.Entry(inputs_stock, width=20)
entry_start.grid(row=2, column=1, sticky="we", pady=4)

ttk.Label(inputs_stock, text="結束日期 (YYYY-MM-DD)：").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
entry_end = ttk.Entry(inputs_stock, width=20)
entry_end.grid(row=3, column=1, sticky="we", pady=4)

inputs_stock.columnconfigure(1, weight=1)

ttk.Button(left_stock, text="查詢", style="Accent.TButton", command=show_stock).pack(pady=(0, 15))

result_container_stock = ttk.LabelFrame(left_stock, text="績效指標", padding=10)
result_container_stock.pack(fill="both", expand=True)

txt_metrics_stock = tk.Text(
    result_container_stock,
    width=50,
    height=6,
    wrap="word",
    font=("Microsoft JhengHei", 11),
    relief="solid",
    bd=1
)
txt_metrics_stock.pack(fill="both", expand=True)
txt_metrics_stock.config(state="disabled")

info_stock_frame = ttk.Frame(result_container_stock)
info_stock_frame.pack(fill="x", pady=(5, 0))
ttk.Button(info_stock_frame, text="ℹ 累積報酬率", style="link.TButton", command=show_stock_cum_info).pack(side="left", padx=(0, 8))
ttk.Button(info_stock_frame, text="ℹ 年化報酬率", style="link.TButton", command=show_stock_cagr_info).pack(side="left", padx=(0, 8))
ttk.Button(info_stock_frame, text="ℹ 年化標準差", style="link.TButton", command=show_stock_std_info).pack(side="left", padx=(0, 8))
ttk.Button(info_stock_frame, text="ℹ Sharpe 比率", style="link.TButton", command=show_stock_sharpe_info).pack(side="left", padx=(0, 8))
ttk.Button(info_stock_frame, text="ℹ Sortino 比率", style="link.TButton", command=show_stock_sortino_info).pack(side="left", padx=(0, 8))

# 右側：分頁 (左右分頁的 Notebook)
right_nb_stock = ttk.Notebook(pw_stock)
pw_stock.add(right_nb_stock, weight=2)

# 「股價走勢」分頁
tab_stock_chart = ttk.Frame(right_nb_stock, padding=5)
right_nb_stock.add(tab_stock_chart, text="股價走勢")

fig_stock = plt.Figure(figsize=(6, 4), dpi=100)
ax_stock = fig_stock.add_subplot(111)
canvas_stock = FigureCanvasTkAgg(fig_stock, master=tab_stock_chart)
canvas_stock.get_tk_widget().pack(fill="both", expand=True)

# 「資料表」分頁（可後續放表格或其他內容）
tab_stock_table = ttk.Frame(right_nb_stock, padding=5)
right_nb_stock.add(tab_stock_table, text="資料表")

# 範例：暫時放一個 Label，未來可換成 Tableview
ttk.Label(tab_stock_table, text="尚未實作資料表").pack(padx=10, pady=10)

stock_query_frame.pack_forget()

# ───── 「退休試算畫面」UI ─────
retirement_frame = ttk.Frame(root, padding=20)
pw = ttk.PanedWindow(retirement_frame, orient="horizontal")
pw.pack(fill="both", expand=True)

left = ttk.Frame(pw, padding=10)
pw.add(left, weight=1)

# 「退休試算頁面」左上：返回首頁按鈕
ttk.Button(left, text="← 返回首頁", command=go_home).pack(anchor="w", pady=(0, 10))

inputs = ttk.LabelFrame(left, text="輸入參數", padding=10)
inputs.pack(fill="x", pady=(0, 10))

ttk.Label(inputs, text="年支出 (元)：").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
entry_expense = ttk.Entry(inputs, width=20)
entry_expense.grid(row=0, column=1, sticky="we", pady=4)

ttk.Label(inputs, text="股票代號：").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
entry_tic_ret = ttk.Entry(inputs, width=20)
entry_tic_ret.grid(row=1, column=1, sticky="we", pady=4)

ttk.Label(inputs, text="市場：").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
combo_market_ret = ttk.Combobox(
    inputs, values=["美國", "台灣", "日本", "英國"], state="readonly", width=18
)
combo_market_ret.current(0)
combo_market_ret.grid(row=2, column=1, sticky="we", pady=4)

ttk.Label(inputs, text="預估通膨率 (%)：").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
entry_infl = ttk.Entry(inputs, width=20)
entry_infl.insert(0, "2.0")
entry_infl.grid(row=3, column=1, sticky="we", pady=4)

ttk.Label(inputs, text="開始日期 (YYYY-MM-DD)：").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
entry_ret_start = ttk.Entry(inputs, width=20)
entry_ret_start.grid(row=4, column=1, sticky="we", pady=4)

ttk.Label(inputs, text="結束日期 (YYYY-MM-DD)：").grid(row=5, column=0, sticky="w", padx=(0, 8), pady=4)
entry_ret_end = ttk.Entry(inputs, width=20)
entry_ret_end.grid(row=5, column=1, sticky="we", pady=4)

inputs.columnconfigure(1, weight=1)

ttk.Button(left, text="計算退休金", style="Accent.TButton", command=calculate_retirement).pack(pady=(0, 15))

result_box = ttk.LabelFrame(left, text="計算結果", padding=10)
result_box.pack(fill="both", expand=True)

txt_result_ret = tk.Text(
    result_box,
    width=50,
    height=9,
    wrap="word",
    font=("Microsoft JhengHei", 11),
    relief="solid",
    bd=1
)
txt_result_ret.pack(fill="both", expand=True, padx=0, pady=(0, 5))
txt_result_ret.config(state="disabled")

# 「退休試算頁面」右側分頁區
right_nb = ttk.Notebook(pw)
pw.add(right_nb, weight=2)

tab_plot_ret = ttk.Frame(right_nb, padding=5)
right_nb.add(tab_plot_ret, text="股價走勢")
fig_ret = plt.Figure(figsize=(6, 4), dpi=100)
ax_ret = fig_ret.add_subplot(111)
canvas_ret = FigureCanvasTkAgg(fig_ret, master=tab_plot_ret)
canvas_ret.get_tk_widget().pack(fill="both", expand=True)

frame_data = ttk.Frame(right_nb, padding=20)
ttk.Label(frame_data, text="尚未實作資料表").pack()
right_nb.add(frame_data, text="資料表")

retirement_frame.pack_forget()

# ───── 啟動主迴圈 ─────
root.mainloop()

