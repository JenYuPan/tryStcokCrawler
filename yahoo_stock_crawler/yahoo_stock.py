import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

# **設定要查詢的股票代號**
stock_ids = ["2330", "2317", "2454", "2603", "2882", "2886"]

# 設定 User-Agent 避免被封鎖
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
}

# **存儲股票數據的列表**
stock_data = []

# **查詢多檔股票**
print(f"{'股票名稱':<10} {'股號':<6} {'價格':<10} {'漲跌':<10} {'漲跌幅':<10} {'最後更新時間'}")
print("=" * 65)

for stock_id in stock_ids:
    url = f'https://tw.stock.yahoo.com/quote/{stock_id}'

    # 取得網頁內容
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")

    # **取得股票名稱**
    stock_name = soup.select_one('h1[class*="C($c-link-text) Fw(b) Fz(24px) Mend(8px)"]')
    stock_name = stock_name.get_text(strip=True) if stock_name else "股票名稱未找到"

    # **取得股價**
    stock_price = soup.select_one('span[class*="Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c)"]')
    price = stock_price.get_text(strip=True) if stock_price else "股價未找到"

    # **初始化變數**
    change_value = "漲跌未找到"
    change_percent_value = "漲跌幅未找到"
    trend_symbol = ""

    # **從 <span class="Mend(4px) Bds(s)"/> 找出 `style` 屬性的 `border-color`**
    trend_arrow = soup.select_one('span.Mend\(4px\).Bds\(s\)')

    if trend_arrow and 'style' in trend_arrow.attrs:
        style = trend_arrow.attrs['style']

        if "border-color:#00ab5e" in style:  # 綠色（下跌）
            trend_symbol = "-"
        elif "border-color:#ff333a" in style:  # 紅色（上漲）
            trend_symbol = "+"
        else:
            trend_symbol="+"

    # **取得漲跌數值**
    change_price = soup.select_one('span[class*="Fz(20px) Fw(b) Lh(1.2) Mend(4px) D(f) Ai(c)"]')
    if change_price:
        change_value = change_price.get_text(strip=True)

    # **取得漲跌幅**
    change_percent = soup.select_one('span[class*="Jc(fe) Fz(20px) Lh(1.2) Fw(b) D(f) Ai(c)"]')
    if change_percent:
        change_percent_value = change_percent.get_text(strip=True)

    # **確保變動數值帶有 `+` 或 `-` 符號**
    if change_value not in ["漲跌未找到", "", None]:
        change_value = trend_symbol + change_value

    if change_percent_value not in ["漲跌幅未找到", "", None]:
        change_percent_value = trend_symbol + change_percent_value

    # **取得最後更新時間**
    last_update = soup.select_one('time')
    last_update_time = last_update.get_text(strip=True) if last_update else "更新時間未找到"

    # **儲存數據到列表**
    stock_data.append([stock_name, stock_id, price, change_value, change_percent_value, last_update_time])

    # **輸出結果**
    print(
        f"{stock_name:<10} {stock_id:<6} {price:<10} {change_value:<10} {change_percent_value:<10} {last_update_time}")

# **轉換為 DataFrame**
df = pd.DataFrame(stock_data, columns=["股票名稱", "股號", "價格", "漲跌", "漲跌幅", "最後更新時間"])

# **Excel 檔案名稱**
excel_file = "yahoo_stock_data.xlsx"

# **檢查 Excel 檔案是否存在**
if os.path.exists(excel_file):
    # **如果 Excel 已存在，則追加數據**
    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df.to_excel(writer, index=False, header=False, startrow=writer.sheets["Sheet1"].max_row)
    print("\n✅ 股票數據已成功追加到 yahoo_stock_data.xlsx！")
else:
    # **如果 Excel 不存在，則創建新檔案**
    df.to_excel(excel_file, index=False, engine="openpyxl")
    print("\n✅ 股票數據已成功寫入 yahoo_stock_data.xlsx！")
