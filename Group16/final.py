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
from ttkbootstrap.tableview import Tableview  # 用來顯示可分頁、可排序的表格
from datetime import datetime

# ───── 初始化 Pygame Mixer ─────
pygame.mixer.init()

# ───── Matplotlib 中文與負號設定 ─────
matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
matplotlib.rcParams["axes.unicode_minus"] = False

# ───── 全域參數 ─────
risk_free_rate = 0      # 無風險利率預設為 0
is_music_playing = False

# ───── 共用工具函式 ─────
def format_ticker(raw: str, market: str) -> str:
    """
    格式化使用者輸入的股票代號，根據市場加上不同後綴。
    例如：raw="0050", market="台灣" → "0050.TW"
    """
    raw = raw.strip().upper()
    if market == "美國":
        return raw
    return raw + {"台灣": ".TW", "日本": ".T", "英國": ".L"}.get(market, "")

def flatten(df: pd.DataFrame) -> pd.DataFrame:
    """
    如果 DataFrame 的欄位是 MultiIndex，就把它扁平化到單層索引。
    """
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df

# ───── 抓取股票資料並計算績效指標 ─────
def fetch_price_and_metrics(ticker: str, start_date=None, end_date=None):
    """
    利用 yfinance 抓取某支股票的歷史資料（Adj Close），計算：
    - 每日報酬(Return)與累積報酬(CumReturn)
    - CAGR（年化報酬率）、年化波動度(std)、Sharpe、Sortino 等指標
    然後把 DataFrame 和這些指標包成 tuple 回傳。
    """
    raw_df = yf.download(ticker, period="max", progress=False, auto_adjust=False)
    df = flatten(raw_df)
    if df.empty or "Adj Close" not in df.columns:
        raise ValueError("抓不到資料或缺少 Adjusted Close 欄位")
    df = df.reset_index()

    # 如果使用者有傳 start_date / end_date，就先篩選
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

    # 計算每日報酬與累積報酬
    df["Return"] = df["Adj Close"].pct_change()
    df.dropna(subset=["Return"], inplace=True)
    df["CumReturn"] = (1 + df["Return"]).cumprod() - 1

    # 取第一筆、最後一筆日期計算總年數
    t0, t1 = df["Date"].iloc[0], df["Date"].iloc[-1]
    years = (t1 - t0).days / 365.25

    # CAGR = (終值 / 起始值)^(1/年數) - 1
    cagr = (df["Adj Close"].iloc[-1] / df["Adj Close"].iloc[0]) ** (1/years) - 1

    # 年化波動度 = 日報酬 std * sqrt(252)
    std_annual = df["Return"].std() * np.sqrt(252)

    # Sharpe = (CAGR - 無風險利率) / 年化波動度
    sharpe = (cagr - risk_free_rate) / std_annual

    # Sortino：只計算負報酬的標準差
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
    """
    按下「查詢」按鈕後執行：
    1. 讀取 entry_stock, combo_market_stock, entry_start, entry_end
    2. 呼叫 fetch_price_and_metrics，取得 df 與指標
    3. 在右側的 matplotlib 畫出累積報酬率曲線
    4. 把指標以文字顯示在 txt_metrics_stock
    """
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

    # 畫累積報酬率 (%)
    ax_stock.clear()
    ax_stock.plot(df["Date"], df["CumReturn"] * 100, color="#1f77b4")
    ax_stock.set_title(f"{ticker} 累積報酬率 (%)")
    ax_stock.set_xlabel("日期")
    ax_stock.set_ylabel("累積報酬率 (%)")
    ax_stock.grid(True)
    fig_stock.autofmt_xdate()
    canvas_stock.draw()

    # 顯示指標
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
    """
    按下「計算退休金」按鈕後執行：
    1. 讀取使用者輸入：年支出、通膨率、股票代號、市場、開始/結束日期
    2. 呼叫 fetch_price_and_metrics，取得 df 與指標
    3. 計算「年度結束日、年化報酬、累積報酬」，並把每年一列插入 table_ret
    4. 把文字結果顯示在 txt_result_ret
    """
    # 1. 讀取輸入
    try:
        expense = float(entry_expense.get())
    except ValueError:
        messagebox.showerror("錯誤", "年支出必須為數字")
        return

    try:
        infl = float(entry_infl.get()) / 100
    except ValueError:
        messagebox.showerror("錯誤", "通膨率必須為數字")
        return

    raw = entry_tic_ret.get().strip()
    if not raw:
        messagebox.showerror("錯誤", "請輸入股票代號")
        return
    ticker = format_ticker(raw, combo_market_ret.get())

    start_date = entry_ret_start.get().strip()
    end_date = entry_ret_end.get().strip()

    # 2. 抓資料並計算指標
    try:
        df, m = fetch_price_and_metrics(
            ticker,
            start_date if start_date else None,
            end_date if end_date else None
        )
    except Exception as e:
        messagebox.showerror("下載失敗", str(e))
        return

    # 如果年化報酬率扣掉通膨後 ≤ 0，就沒辦法做安全提領
    real_withdraw = (m["ann"] / 100) - infl
    if real_withdraw <= 0:
        messagebox.showwarning("無法計算", "年化報酬不足以支付通膨率")
        return
    need_capital = expense / real_withdraw

    # 3. 在右側分頁的「股價走勢」畫調整後收盤價
    ax_ret.clear()
    ax_ret.plot(df["Date"], df["Adj Close"], color="#ff7f0e")
    ax_ret.set_title(f"{ticker} 調整後收盤價")
    ax_ret.set_xlabel("日期")
    ax_ret.set_ylabel("價格")
    ax_ret.grid(True)
    fig_ret.autofmt_xdate()
    canvas_ret.draw()

    # 4. 填入年度明細到 table_ret
    table_ret.clear()
    df_copy = df.copy()
    df_copy["Year"] = df_copy["Date"].dt.year
    summary = df_copy.groupby("Year").apply(lambda g: pd.Series({
        "end_date": g["Date"].max().date(),
        "end_price": g.loc[g["Date"].idxmax(), "Adj Close"],
        "annual_return": (g.loc[g["Date"].idxmax(), "Adj Close"] / g.loc[g["Date"].idxmin(), "Adj Close"] - 1) * 100,
        "cum_return": g["CumReturn"].max() * 100
    })).reset_index()

    for _, row in summary.iterrows():
        table_ret.insert_row([
            row["Year"],
            row["end_date"],
            f"{row['end_price']:.2f}",
            f"{row['annual_return']:.2f}",
            f"{row['cum_return']:.2f}"
        ])

    # 5. 把文字結果顯示出來
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

# 使用背景圖並模糊
original_bg = Image.open("background4.jpg").filter(ImageFilter.GaussianBlur(radius=2))
bg_pre = original_bg.resize((1280, 800))
bg_photo = ImageTk.PhotoImage(bg_pre)
main_canvas.bg_photo = bg_photo
bg_image_id = main_canvas.create_image(0, 0, image=bg_photo, anchor="nw")

# 「投資助手」標題（含陰影效果）
initial_title_font = ("Noto Sans TC", 36, "bold")
title_shadow = main_canvas.create_text(
    0, 0, text="投資助手", font=initial_title_font, fill="#484848", anchor="center"
)
title_id = main_canvas.create_text(
    0, 0, text="投資助手", font=initial_title_font, fill="#ffffff", anchor="center"
)

# 首頁按鈕：查詢股票＆退休試算
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

# 音樂開關按鈕
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

# 當視窗大小改變時，動態調整背景與文字大小
def update_layout(event=None):
    w = main_canvas.winfo_width()
    h = main_canvas.winfo_height()
    if w < 100 or h < 100:
        return

    try:
        # 重設背景圖大小
        resized = original_bg.resize((w, h))
        new_bg = ImageTk.PhotoImage(resized)
        main_canvas.itemconfig(bg_image_id, image=new_bg)
        main_canvas.bg_photo = new_bg

        # 調整標題字型大小
        title_size = max(24, min(48, int(h * 0.06)))
        dyn_title_font = ("Noto Sans TC", title_size, "bold")
        main_canvas.itemconfig(title_shadow, font=dyn_title_font)
        main_canvas.itemconfig(title_id, font=dyn_title_font)

        # 調整按鈕字型大小與間距
        btn_font_size = max(10, int(h * 0.025))
        dyn_btn_font = ("Noto Sans TC", btn_font_size)
        btn_pad_y = int(h * 0.015)
        style.configure("Dyn.TButton", font=dyn_btn_font, padding=(20, btn_pad_y))
        btn1.configure(style="Dyn.TButton")
        btn2.configure(style="Dyn.TButton")

        # 重新定位
        main_canvas.coords(title_shadow, w // 2 + 3, int(h * 0.12))
        main_canvas.coords(title_id, w // 2, int(h * 0.11))
        main_canvas.coords(btn_window_id, w // 2, int(h * 0.25))
        main_canvas.coords(speaker_id, w - 30, 30)

    except Exception as e:
        print("resize error:", e)

main_canvas.bind("<Configure>", update_layout)

# ───── 「股票查詢畫面」UI ─────
stock_query_frame = ttk.Frame(root, padding=20)

left_stock = ttk.Frame(stock_query_frame, padding=10)
left_stock.pack(side="left", fill="y")

# 返回首頁
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

right_stock = ttk.Frame(stock_query_frame, padding=10)
right_stock.pack(side="right", fill="both", expand=True)

fig_stock = plt.Figure(figsize=(6, 4), dpi=100)
ax_stock = fig_stock.add_subplot(111)
canvas_stock = FigureCanvasTkAgg(fig_stock, master=right_stock)
canvas_stock.get_tk_widget().pack(fill="both", expand=True)

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

# ───── 「退休試算頁面」右側分頁區 ─────
right_nb = ttk.Notebook(pw)
pw.add(right_nb, weight=2)

tab_plot_ret = ttk.Frame(right_nb, padding=5)
right_nb.add(tab_plot_ret, text="股價走勢")
fig_ret = plt.Figure(figsize=(6, 4), dpi=100)
ax_ret = fig_ret.add_subplot(111)
canvas_ret = FigureCanvasTkAgg(fig_ret, master=tab_plot_ret)
canvas_ret.get_tk_widget().pack(fill="both", expand=True)

# 把「尚未實作資料表」改成 Tableview
frame_data = ttk.Frame(right_nb, padding=5)
right_nb.add(frame_data, text="資料表")

table_ret = Tableview(frame_data,
                      coldata=[
                          {"text": "年份", "stretch": True},
                          {"text": "結束日期", "stretch": True},
                          {"text": "收盤價", "stretch": True},
                          {"text": "年報酬率 (%)", "stretch": True},
                          {"text": "累積報酬率 (%)", "stretch": True}
                      ],
                      rowdata=[],
                      paginated=True,
                      height=15)
table_ret.pack(fill="both", expand=True, padx=10, pady=10)

retirement_frame.pack_forget()

# ───── 啟動主迴圈 ─────
root.mainloop()

