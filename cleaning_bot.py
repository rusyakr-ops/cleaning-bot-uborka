# -*- coding: utf-8 -*-
import os
import time
import datetime
import pandas as pd
import requests
import schedule

# ==== НАСТРОЙКИ ==== #
# Эти две штуки возьмём из переменных окружения на Render
TOKEN = os.environ["8522306269:AAEz4k3HKuwQabTbJgUit1HsM7YEESS7Og4"]          # токен бота
CHAT_ID = int(os.environ["-1003483287470"]) # chat_id группы (число, со знаком -)
CLEANING_TIME = "17:00"                       # время уборки, как в сообщении
EXCEL_PATH = "Uborka.xlsx"                    # файл с графиком
# ==================== #

ZONE_DETAILS = {
    "Полы": "🤸 Подмести и помыть пол на кухне (включая труднодоступные места: возле дивана, под полкой обуви, под столом — *поднять стулья*)",
    "Поверхности": "🧽 Вытереть стол, помыть плиту и холодильник (со стороны плиты), помыть подставку и раковину, разложить посуду, протереть столешницу и диван",
    "Туалет": "🚽 Вытереть крышку, подмести и помыть пол, залить средство, убрать все лишнее",
    "Ванна": "🛁 Помыть раковину (убрать баночки), зеркало, помыть пол (в т.ч. под раковиной), убрать волосы из слива при необходимости",
}

def get_tasks_for_nearest_date(target_date: datetime.date, df: pd.DataFrame):
    header_row_idx = None
    for i in range(len(df)):
        if str(df.iloc[i, 0]).strip() == "Имя/Зона":
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Не найден заголовок 'Имя/Зона'.")

    date_row_idx = header_row_idx - 1
    date_cols = []
    for col in range(1, df.shape[1]):
        val = df.iloc[date_row_idx, col]
        if isinstance(val, (datetime.date, datetime.datetime, pd.Timestamp)):
            date_cols.append((col, pd.to_datetime(val).date()))

    if not date_cols:
        raise ValueError("Не найдено ни одной даты в строке с датами.")

    candidates = [(c, d) for c, d in date_cols if d >= target_date]
    if candidates:
        target_col, chosen_date = min(candidates, key=lambda x: x[1])
    else:
        target_col, chosen_date = max(date_cols, key=lambda x: x[1])

    date_cols_sorted = sorted(date_cols, key=lambda x: x[0])
    idx = [c for c, _ in date_cols_sorted].index(target_col)

    if idx < len(date_cols_sorted) - 1:
        next_col = date_cols_sorted[idx + 1][0]
        group_cols = list(range(target_col, next_col))
    else:
        group_cols = list(range(target_col, df.shape[1]))

    name_rows = []
    r = header_row_idx + 1
    while r < len(df):
        val = df.iloc[r, 0]
        if pd.isna(val) or str(val).strip() == "":
            break
        name_rows.append(r)
        r += 1

    tasks = {}
    for r in name_rows:
        name = str(df.iloc[r, 0]).strip()
        zones = []
        for c in group_cols:
            cell = df.iloc[r, c]
            if isinstance(cell, str) and cell.strip().lower() in ["x", "х"]:
                zones.append(str(df.iloc[header_row_idx, c]).strip())
        tasks[name] = zones

    return chosen_date, tasks

def build_message(chosen_date, tasks, cleaning_time):
    date_str = chosen_date.strftime('%d.%m.%Y')
    lines = [
        f"🧹 <b>Сегодня уборка ({date_str}) в {cleaning_time}!</b>\n",
        "✨ <b>Обязанности:</b>\n"
    ]
    for name, zones in tasks.items():
        if zones:
            lines.append(f"<b>{name}</b>:")
            for z in zones:
                detail = ZONE_DETAILS.get(z, f"▸ {z}")
                lines.append(f" ▸ {detail}")
            lines.append("")
    lines.append("💧 Если закончили раньше — отметьтесь в чате 😉")
    lines.append("🫧 Хорошего настроения и чистоты!")
    return "\n".join(lines).strip()

def send_message(text):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    data = {"chat_id": CHAT_ID, "text": text, "parse_mode": "HTML"}
    resp = requests.post(url, data=data)
    print("Статус отправки:", resp.status_code, resp.text)

def send_cleaning_message():
    df = pd.read_excel(EXCEL_PATH, header=None)
    today = datetime.date.today()
    chosen_date, tasks = get_tasks_for_nearest_date(today, df)
    msg = build_message(chosen_date, tasks, CLEANING_TIME)
    print("Сообщение:")
    print(msg)
    send_message(msg)

if __name__ == "__main__":
    # Render работает в UTC. Таллин = UTC+2 зимой.
    # 11:00 по Таллину -> 09:00 по UTC
    schedule.every().sunday.at("09:00").do(send_cleaning_message)

    print("Запущен планировщик. Жду расписание...")
    while True:
        schedule.run_pending()
        time.sleep(60)
