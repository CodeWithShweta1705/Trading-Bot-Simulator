import tkinter as tk
from tkinter import messagebox, ttk
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- GLOBAL DATA STORAGE ---
current_analysis_data = pd.DataFrame() 
current_symbol = ""

# --- CORE FUNCTIONS ---
def get_analysis():
    global current_analysis_data, current_symbol
    symbol = entry_symbol.get().upper().strip()
    
    if not symbol:
        messagebox.showwarning("Input Error", "Please enter a Stock Symbol!")
        return
        
    if not symbol.endswith(".NS"): 
        symbol += ".NS"
    
    current_symbol = symbol
    
    period_map = {
        "1 Month": "1mo", 
        "3 Months": "3mo", 
        "6 Months": "6mo", 
        "1 Year": "1y", 
        "5 Years": "5y"
    }
    selected_period = period_map[combo_period.get()]

    try:
        # 1. Fetching Data
        stock = yf.Ticker(symbol)
        df = stock.history(period=selected_period)
        
        if df.empty:
            messagebox.showwarning("Data Error", "No data found. Please check the symbol (e.g. RELIANCE).")
            return

        # Data store for Excel (Fixing Global Data Issue)
        current_analysis_data = df.copy()

        # 2. RSI Calculation for Buy/Sell Signals
        delta = df['Close'].diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs))
        last_rsi = rsi.iloc[-1]

        # 3. Update UI Labels
        current_price = df['Close'].iloc[-1]
        lbl_price.config(text=f"Current Price: ₹{current_price:.2f}", fg="#2ecc71")
        
        if last_rsi < 35:
            lbl_signal.config(text="SIGNAL: BUY (Oversold)", fg="#2ecc71")
        elif last_rsi > 65:
            lbl_signal.config(text="SIGNAL: SELL (Overbought)", fg="#e74c3c")
        else:
            lbl_signal.config(text="SIGNAL: HOLD (Neutral)", fg="#f1c40f")

        # 4. Refreshing the Chart
        for widget in chart_frame.winfo_children():
            widget.destroy()
            
        fig, ax = plt.subplots(figsize=(5, 3), dpi=90)
        fig.patch.set_facecolor('#121212')
        ax.set_facecolor('#1e1e1e')
        ax.plot(df.index, df['Close'], color='#f1c40f', linewidth=2)
        ax.tick_params(colors='white', labelsize=8)
        ax.grid(color='#333333', linestyle='--')
        plt.xticks(rotation=25)
        
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch data: {e}")

def export_data():
    global current_analysis_data, current_symbol
    if not current_analysis_data.empty:
        try:
            # --- IMPORTANT: Timezone Fix for Excel ---
            data_to_save = current_analysis_data.copy()
            if data_to_save.index.tz is not None:
                data_to_save.index = data_to_save.index.tz_localize(None)
            
            excel_name = f"{current_symbol}_Report.xlsx"
            csv_name = f"{current_symbol}_Data.csv"
            
            # Saving in both formats
            data_to_save.to_excel(excel_name)
            data_to_save.to_csv(csv_name)
            
            messagebox.showinfo("Success", f"Reports Saved!\n1. {excel_name}\n2. {csv_name}\n\nCheck your folder now!")
        except Exception as e:
            messagebox.showerror("Error", f"Excel Save Failed: {e}\nTry installing: pip install openpyxl")
    else:
        messagebox.showwarning("Warning", "Please RUN ANALYSIS first!")

# --- MAIN UI DESIGN ---
root = tk.Tk()
root.title("Shweta's Trading Intelligence")
root.geometry("600x750")
root.configure(bg="#121212")

# Title Header
tk.Label(root, text="SHWETA'S TRADING INTELLIGENCE", font=("Segoe UI", 18, "bold"), bg="#121212", fg="#f1c40f").pack(pady=20)

# Input Row (Symbol + Period)
input_frame = tk.Frame(root, bg="#121212")
input_frame.pack(pady=10)

tk.Label(input_frame, text="Stock Symbol:", bg="#121212", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
entry_symbol = tk.Entry(input_frame, font=("Arial", 12), width=12, justify='center')
entry_symbol.insert(0, "RELIANCE")
entry_symbol.pack(side=tk.LEFT, padx=5)

combo_period = ttk.Combobox(input_frame, values=["1 Month", "3 Months", "6 Months", "1 Year", "5 Years"], state="readonly", width=12)
combo_period.current(0)
combo_period.pack(side=tk.LEFT, padx=5)

# Primary Button
tk.Button(root, text="RUN ANALYSIS", command=get_analysis, bg="#2980b9", fg="white", font=("Arial", 11, "bold"), width=25, cursor="hand2").pack(pady=15)

# Price Display
lbl_price = tk.Label(root, text="Price: --", font=("Segoe UI", 16), bg="#121212", fg="#2ecc71")
lbl_price.pack()

# Signal Display
lbl_signal = tk.Label(root, text="SIGNAL: --", font=("Segoe UI", 20, "bold"), bg="#121212", fg="#f1c40f")
lbl_signal.pack(pady=10)

# Chart Frame Area
chart_frame = tk.Frame(root, bg="#1e1e1e", bd=2, relief="groove")
chart_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

# Export Button at Bottom
tk.Button(root, text="EXPORT REPORT (EXCEL & CSV)", command=export_data, bg="#27ae60", fg="white", font=("Arial", 12, "bold"), width=35, cursor="hand2").pack(pady=25)

# System Footer
tk.Label(root, text="Status: Ready to Analyze", font=("Arial", 8), bg="#121212", fg="#555555").pack(side=tk.BOTTOM, pady=5)

root.mainloop()