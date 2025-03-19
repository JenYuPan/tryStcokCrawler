import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

#股票號碼
stock_ids = ["2330", "2317", "2454", "2603", "2882", "2886"]

# User-Agent
headers = {
    "User-Agent": "Your User Agent"
}

# 股票的陣列
stock_data = []

# 定義欄位名稱
columns = ["股票名稱", "股號", "價格", "漲跌", "漲跌幅", "昨日收盤", "總量", "昨量", "開盤價", "最高價", "均價", "最低價", "最後更新時間"]

print(f"{'股票名稱':<10} {'股號':<6} {'價格':<10} {'漲跌':<10} {'漲跌幅':<10} {'昨日收盤':<10} {'總量':<10} {'昨量':<10} {'開盤價':<10} {'最高價':<10} {'均價':<10} {'最低價':<10} {'最後更新時間'}")
print("=" * 150)

for stock_id in stock_ids:
    url = f'https://tw.stock.yahoo.com/quote/{stock_id}'
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    stock_info = {col: "0" for col in columns}
    stock_info["股號"] = stock_id
    # 股票名稱
    stock_name = soup.select_one('h1[class*="C($c-link-text) Fw(b) Fz(24px) Mend(8px)"]')
    if stock_name:
        stock_info["股票名稱"] = stock_name.get_text(strip=True)

    # 股價
    stock_price = soup.select_one('span[class*="Fz(32px) Fw(b) Lh(1) Mend(16px) D(f) Ai(c)"]')
    if stock_price:
        stock_info["價格"] = stock_price.get_text(strip=True)

    # 符號
    trend_symbol = ""
    trend_arrow = soup.select_one('span.Mend\(4px\).Bds\(s\)')
    if trend_arrow and 'style' in trend_arrow.attrs:
        style = trend_arrow.attrs['style']
        # 綠色（下跌）
        if "border-color:#00ab5e" in style:
            trend_symbol = "-"
        # 紅色（上漲）
        elif "border-color:#ff333a" in style:
            trend_symbol = "+"

    price_details = soup.select("li.price-detail-item")

    #數據變量
    total_volume = "0"
    for item in price_details:
        spans = item.find_all("span")

        if len(spans) >= 2:
            label = spans[0].get_text(strip=True)   #第一個span是標籤
            value = spans[1].get_text(strip=True)   #第二個span是數值

            if label == "昨收":
                stock_info["昨日收盤"] = value
            elif label == "總量":
                total_volume = value
            elif label == "成交量":
                if total_volume == "0":
                    total_volume = value
            elif label == "昨量":
                stock_info["昨量"] = value
            elif label == "開盤":
                stock_info["開盤價"] = value
            elif label == "最高":
                stock_info["最高價"] = value
            elif label == "均價":
                stock_info["均價"] = value
            elif label == "最低":
                stock_info["最低價"] = value
            elif label == "漲跌":
                stock_info["漲跌"] = trend_symbol + value
            elif label == "漲跌幅":
                stock_info["漲跌幅"] = trend_symbol + value

    #總量判斷
    stock_info["總量"] = total_volume if total_volume != "0" else "未找到"

    #最後更新時間
    last_update = soup.select_one('time')
    if last_update:
        stock_info["最後更新時間"] = last_update.get_text(strip=True)

    #存入列表
    stock_data.append([stock_info[col] for col in columns])

    #輸出結果
    print(f"{stock_info['股票名稱']:<10} {stock_info['股號']:<6} {stock_info['價格']:<10} {stock_info['漲跌']:<10} {stock_info['漲跌幅']:<10} {stock_info['昨日收盤']:<10} "
          f"{stock_info['總量']:<10} {stock_info['昨量']:<10} {stock_info['開盤價']:<10} {stock_info['最高價']:<10} {stock_info['均價']:<10} {stock_info['最低價']:<10} {stock_info['最後更新時間']}")

# 轉換為格式
df = pd.DataFrame(stock_data, columns=columns)
excel_file = "yahoo_stock_data.xlsx"

# 檢查檔案是否存在
if os.path.exists(excel_file):
    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df.to_excel(writer, index=False, header=False, startrow=writer.sheets["Sheet1"].max_row)
    print("\n✅ 股票數據已成功追加到 yahoo_stock_data.xlsx！")
else:
    df.to_excel(excel_file, index=False, engine="openpyxl")
    print("\n✅ 股票數據已成功寫入 yahoo_stock_data.xlsx！")
