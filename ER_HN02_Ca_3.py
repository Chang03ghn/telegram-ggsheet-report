import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import dataframe_image as dfi
import telegram
import telegram.request
import asyncio

# =========== CẤU HÌNH ============
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1x982zCBXltNhsIlvFpp516Pt1F7YGDPcJ9oUzKbwxtY/edit?pli=1&gid=1208018138#gid=1208018138'
CREDENTIALS_FILE = 'mybot-468103-2c8133fa8479.json'
TELEGRAM_BOT_TOKEN = '8363729641:AAH4zpLvUKmSIYajkRxw_JZIe_L1HqSkvhw'
TELEGRAM_CHAT_ID = -4918308974
IMAGE_FILE = 'image_sheet3.png'

# =========== LẤY DỮ LIỆU GOOGLE SHEET ============
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

sheet = client.open_by_url(GOOGLE_SHEET_URL).worksheet("ER-HN02-Ca 3")
data = sheet.get_all_values()
df = pd.DataFrame(data)

# =========== STYLING ============
def highlight_format(x):
    df_style = pd.DataFrame('', index=x.index, columns=x.columns)

    # Dòng 4 (index 3): cam nhạt
    df_style.iloc[3, :] = 'background-color: #ffe599'

    # Dòng 5, 18 (index 4, 17): vàng nhạt
    for i in [4, 17]:
        df_style.iloc[i, :] = 'background-color: #fff2cc'

    # Dòng 7,10,13,22,25,28 (index): nâu nhạt
    for i in [6, 9, 12,15, 21, 24, 27,30]:
        df_style.iloc[i, :] = 'background-color: #d9d2e9'

    # Dòng 16, 31 (index 15, 30): xanh lá nhạt
    for i in [16, 31]:
        df_style.iloc[i, :] = 'background-color: #b6d7a8'

    # Dòng cuối cùng: xanh nhạt
    df_style.iloc[-1, :] = 'background-color: #d9ead3'

    # Cột 4 (index 3): vàng nhạt
    df_style.iloc[:, 2] = df_style.iloc[:, 3].where(df_style.iloc[:, 3] != '', 'background-color: #fff2cc')

    # 3 cột đầu (index 0-2): xám nhạt
    for col in range(3):
        df_style.iloc[:, col] = df_style.iloc[:, col].where(df_style.iloc[:, col] != '', 'background-color: #f0f0f0')

    return df_style


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

# =========== XUẤT ẢNH ============
dfi.export(styled_df, IMAGE_FILE, dpi=150)  # Tăng dpi nếu ảnh bị mờ

# =========== GỬI ẢNH TỚI TELEGRAM ============
async def send_photo_to_telegram():
    request = telegram.request.HTTPXRequest(
        connect_timeout=30.0,
        read_timeout=60.0
    )
    bot = telegram.Bot(token=TELEGRAM_BOT_TOKEN, request=request)
    with open(IMAGE_FILE, 'rb') as f:
        await bot.send_photo(chat_id=TELEGRAM_CHAT_ID, photo=f, caption="📊 Báo cáo Kho HN02 - CA 3")

# Gọi hàm gửi
asyncio.run(send_photo_to_telegram())

print("✅ Gửi ảnh định dạng đẹp (tách dòng/cột rõ ràng) thành công!")
