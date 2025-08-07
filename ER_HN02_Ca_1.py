import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import dataframe_image as dfi
import telegram
from datetime import datetime
import os
import json

# =========== CẤU HÌNH ============

GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1x982zCBXltNhsIlvFpp516Pt1F7YGDPcJ9oUzKbwxtY'

# Load credentials từ biến môi trường
credentials_json = os.environ['GOOGLE_CREDENTIALS']
with open('creds.json', 'w') as f:
    f.write(credentials_json)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
client = gspread.authorize(creds)

TELEGRAM_BOT_TOKEN = os.environ['TELEGRAM_BOT_TOKEN']
TELEGRAM_CHAT_ID_1 = -4901592432
TELEGRAM_CHAT_ID_2 = -4901592432

# =========== HÀM XỬ LÝ CỘT TRÙNG ============
def make_columns_unique(columns):
    seen = {}
    result = []
    for col in columns:
        if col not in seen:
            seen[col] = 0
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")
    return result

# =========== STYLING ============

def highlight_format(x):
    df_style = pd.DataFrame('', index=x.index, columns=x.columns)

    # Dòng 4 (index 3): cam nhạt
    df_style.iloc[3, :] = 'background-color: #ffe599'

    # Cột C từ dòng 4 trở đi (index 2)
    df_style.iloc[4:, 2] = 'background-color: #fff2cc'

    # Cột A (index 0): chứa 'Total' => tô xanh, còn lại tô vàng nhạt
    for i in range(len(x)):
        val = str(x.iloc[i, 0]).lower()
        if 'total' in val:
            df_style.iloc[i, 0] = 'background-color: #d9ead3'
        elif val.strip() != '':
            df_style.iloc[i, 0] = 'background-color: #fff2cc'

    # Cột B (index 1): nếu chứa 'Total' => tô cả dòng tím nhạt
    for i in range(len(x)):
        val_b = str(x.iloc[i, 1]).lower()
        if 'total' in val_b:
            df_style.iloc[i, :] = 'background-color: #d9d2e9'

    # Dòng cuối cùng: xanh nhạt
    df_style.iloc[-1, :] = 'background-color: #d9ead3'

    return df_style

# =========== XUẤT ẢNH & GỬI TELEGRAM ============

def export_and_send(sheet_name, image_name, chat_id, caption):
    sheet = client.open_by_url(GOOGLE_SHEET_URL).worksheet(sheet_name)
    data = sheet.get_all_values()

    # Xử lý dataframe
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = make_columns_unique(df.columns)  # ✅ chống trùng tên cột
    df.reset_index(drop=True, inplace=True)

    # Tô màu và xuất ảnh
    styled_df = df.style \
        .set_properties(**{
            'text-align': 'center',
            'border': '1px solid black',
            'font-size': '10pt',
            'font-weight': 'bold'
        }) \
        .set_table_styles([{
            'selector': 'th',
            'props': [('background-color', '#004c99'), ('color', 'white'), ('font-weight', 'bold')]
        }]) \
        .apply(highlight_format, axis=None)

    dfi.export(styled_df, image_name, dpi=150)

    # Gửi Telegram
    bot = telegram.Bot(token=TELEGRAM_BOT_TOKEN)
    with open(image_name, 'rb') as f:
        bot.send_photo(chat_id=chat_id, photo=f, caption=caption)

# =========== MAIN ============

def main():
    now = datetime.now().strftime('%H:%M %d-%m-%Y')
    export_and_send("ER-HN02-Ca 1", "image_ca1.png", TELEGRAM_CHAT_ID_1, f"📊 Báo cáo HN02 - CA 1\n🕐 {now}")
    export_and_send("ER-HY01-Ca 1", "image_ca2.png", TELEGRAM_CHAT_ID_2, f"📊 Báo cáo Hy01 - CA 1\n🕐 {now}")

if __name__ == "__main__":
    main()