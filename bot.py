import asyncio
import sqlite3
import os
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, CallbackQuery, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton, BotCommand
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ========== КОНФИГ ==========
BOT_TOKEN = "8803530037:AAHVuMAb6gIzGXBKH8qbteZtFyttz6_hzh0"
DB_NAME = "trades.db"
BT_DB_NAME = "backtests.db"
TRADES_PER_PAGE = 10

# ========== ЛОКАЛИЗАЦИЯ ==========
TEXTS = {
    "ru": {
        "select_mode": "🎛 **Выберите режим работы:**",
        "mode_real": "📊 Реальная торговля",
        "mode_backtest": "🔄 Бэктест",
        "add_trade": "➕ Сделка",
        "list_trades": "📋 Список сделок",
        "stats": "📊 Статистика",
        "excel_real": "📎 Excel (Реал)",
        "excel_backtest": "📎 Excel (Бэктест)",
        "clear": "🗑 Очистить всё",
        "settings": "⚙️ Настройки",
        "support": "📞 Поддержка",
        "back": "🔙 Назад",
        "yes": "✅ Да",
        "no": "❌ Нет",
        "confirm_clear": "⚠️ ДА, УДАЛИТЬ ВСЕ СДЕЛКИ",
        "confirm_delete": "⚠️ Удалить сделку #{id}?",
        "deleted": "🗑 Сделка #{id} удалена!",
        "cleared": "🗑 Журнал очищен.",
        "no_data": "📭 Нет данных.",
        "no_data_add_trade": "📭 Нет данных. Добавьте первую сделку через /new или меню.",
        "enter_asset": "📝 Введите тикер (BTC, ETH, TON, AAPL):",
        "choose_direction": "📈 Выберите направление:",
        "long": "🟢 LONG",
        "short": "🔴 SHORT",
        "enter_entry_price": "💰 Введите цену входа:",
        "enter_exit_price": "💰 Введите цену выхода:",
        "enter_volume": "📊 Введите объём позиции:",
        "choose_result": "🎯 Как закрылась сделка?",
        "take": "✅ Тейк",
        "stop": "❌ Стоп",
        "bu": "⚖️ БУ",
        "enter_comment": "📝 Введите комментарий (отправьте '-' чтобы пропустить):",
        "add_link_question": "🔗 Хотите добавить ссылку на график?",
        "enter_link": "Отправьте ссылку:",
        "enter_timeframe": "Какой это таймфрейм? (15м, 1ч, 4ч, 1д, 1н, 1м)",
        "link_saved": "✅ Ссылка сохранена! Добавить ещё?",
        "enter_date": "📅 Введите дату (ДД.ММ.ГГГГ) или 'сегодня':",
        "trade_saved": "✅ Сделка сохранена!",
        "enter_emotion": "😊 Какие эмоции были?",
        "emotion_calm": "😊 Спокойствие",
        "emotion_fear": "😨 Страх",
        "emotion_greed": "😈 Жадность",
        "emotion_tilt": "🤬 Тильт",
        "emotion_confidence": "😌 Уверенность",
        "edit_field_select": "✏️ **Редактирование сделки #{id}**\n\nВыберите поле:",
        "edit_asset": "🪙 Актив",
        "edit_direction": "📈 Направление",
        "edit_entry_price": "💰 Цена входа",
        "edit_exit_price": "💰 Цена выхода",
        "edit_volume": "📊 Объём",
        "edit_result": "🎯 Исход",
        "edit_comment": "📝 Комментарий",
        "edit_date": "📅 Дата",
        "edit_emotion": "😊 Эмоции",
        "enter_new_value": "Введите новое значение для {field}:",
        "field_updated": "✅ {field} обновлено!",
        "sort_label": "📅 Сортировка:",
        "sort_newest": "Сначала новые",
        "sort_oldest": "Сначала старые",
        "filter_label": "🔍 Фильтры:",
        "filter_all": "Все",
        "filter_take": "✅ Тейк",
        "filter_stop": "❌ Стоп",
        "filter_bu": "⚖️ БУ",
        "filter_asset": "💰 По активу",
        "filter_date": "📅 По дате",
        "filter_clear": "🗑 Сбросить фильтры",
        "filter_date_day": "За день",
        "filter_date_week": "За неделю",
        "filter_date_month": "За месяц",
        "select_asset_filter": "💰 **Выберите актив:**",
        "select_date_filter": "📅 **Выберите период:**",
        "trade_detail": "📋 **Сделка #{id}**\n\n🪙 Актив: {asset}\n📈 Направление: {direction}\n💰 Вход: ${entry}\n💰 Выход: ${exit}\n📊 Объём: {volume}\n💵 P&L: ${pnl}\n🎯 Исход: {result}\n📅 Дата: {date}\n😊 Эмоции: {emotion}\n🔗 Ссылки:\n{links}\n📝 Комментарий: {comment}",
        "recent_trades": "📋 **Сделки** (страница {page}/{total_pages})",
        "prev": "⬅️ Назад",
        "next": "Вперед ➡️",
        "refresh": "🔄 Обновить",
        "sort": "📅 Сортировка",
        "filter": "🔍 Фильтры",
        "edit": "✏️ Редактировать",
        "delete": "🗑 Удалить",
        "back_to_list": "🔙 К списку",
        "stats_header": "📊 Ваша статистика",
        "bt_period_start": "📅 Введите НАЧАЛО периода (ДД.ММ.ГГГГ):",
        "bt_period_end": "📅 Введите КОНЕЦ периода (ДД.ММ.ГГГГ):",
        "bt_timeframe": "⏱ Введите таймфрейм (M5, H1, H4, D1, W1):",
        "enter_exit_price_bt": "💰 Введите цену выхода:",
        "enter_link_bt": "🔗 Ссылка на скриншот (0 если нет):",
        "error_number": "❌ Ошибка! Введите число.",
        "error_date": "❌ Ошибка! Введите дату в формате ДД.ММ.ГГГГ.",
        "select_language": "🌐 Выберите язык / Choose language:",
        "language_set": "✅ Язык установлен: русский",
        "language_set_en": "✅ Language set: English",
        "support_info": "📞 **Поддержка**\n\nПо вопросам пишите: @ваш_username",
        "change_lang": "🌐 Сменить язык",
        "support": "📞 Поддержка",
        "settings_menu": "⚙️ **Настройки**",
        "stats_menu": "📊 **Меню статистики**\n\nВыберите тип статистики:",
        "stats_all": "📈 Вся статистика",
        "stats_by_asset": "💰 По активам",
        "stats_by_date": "📅 По дате",
        "stats_by_emotion": "😊 По эмоциям",
        "stats_recent": "🕒 Недавние сделки",
        "stats_sort": "🔄 Сортировка (Новые/Старые)",
        "stats_back": "🔙 Назад в меню режима",
        "bt_stats_header": "📊 Статистика бэктеста",
        "bt_total_trades": "📋 Всего бэктестов: {total}",
        "select_asset": "💰 **Выберите актив:**",
        "select_period": "📅 **Выберите период:**",
        "period_day": "📆 День",
        "period_week": "📅 Неделя",
        "period_month": "📊 Месяц"
    },
    "en": {
        "select_mode": "🎛 **Select mode:**",
        "mode_real": "📊 Real Trading",
        "mode_backtest": "🔄 Backtest",
        "add_trade": "➕ Add Trade",
        "list_trades": "📋 Trade List",
        "stats": "📊 Statistics",
        "excel_real": "📎 Excel (Real)",
        "excel_backtest": "📎 Excel (Backtest)",
        "clear": "🗑 Clear All",
        "settings": "⚙️ Settings",
        "support": "📞 Support",
        "back": "🔙 Back",
        "yes": "✅ Yes",
        "no": "❌ No",
        "confirm_clear": "⚠️ YES, DELETE ALL TRADES",
        "confirm_delete": "⚠️ Delete trade #{id}?",
        "deleted": "🗑 Trade #{id} deleted!",
        "cleared": "🗑 Journal cleared.",
        "no_data": "📭 No data.",
        "no_data_add_trade": "📭 No data. Add your first trade via /new.",
        "enter_asset": "📝 Enter ticker (BTC, ETH, TON, AAPL):",
        "choose_direction": "📈 Choose direction:",
        "long": "🟢 LONG",
        "short": "🔴 SHORT",
        "enter_entry_price": "💰 Enter entry price:",
        "enter_exit_price": "💰 Enter exit price:",
        "enter_volume": "📊 Enter position size:",
        "choose_result": "🎯 How did the trade close?",
        "take": "✅ Take",
        "stop": "❌ Stop",
        "bu": "⚖️ BE",
        "enter_comment": "📝 Enter comment (send '-' to skip):",
        "add_link_question": "🔗 Add chart link?",
        "enter_link": "Send the link:",
        "enter_timeframe": "What timeframe? (15m, 1h, 4h, 1d, 1w, 1M)",
        "link_saved": "✅ Link saved! Add another?",
        "enter_date": "📅 Enter date (DD.MM.YYYY) or 'today':",
        "trade_saved": "✅ Trade saved!",
        "enter_emotion": "😊 What emotion did you feel?",
        "emotion_calm": "😊 Calm",
        "emotion_fear": "😨 Fear",
        "emotion_greed": "😈 Greed",
        "emotion_tilt": "🤬 Tilt",
        "emotion_confidence": "😌 Confidence",
        "edit_field_select": "✏️ **Edit trade #{id}**\n\nSelect field:",
        "edit_asset": "🪙 Asset",
        "edit_direction": "📈 Direction",
        "edit_entry_price": "💰 Entry price",
        "edit_exit_price": "💰 Exit price",
        "edit_volume": "📊 Volume",
        "edit_result": "🎯 Outcome",
        "edit_comment": "📝 Comment",
        "edit_date": "📅 Date",
        "edit_emotion": "😊 Emotion",
        "enter_new_value": "Enter new value for {field}:",
        "field_updated": "✅ {field} updated!",
        "sort_label": "📅 Sort:",
        "sort_newest": "Newest first",
        "sort_oldest": "Oldest first",
        "filter_label": "🔍 Filters:",
        "filter_all": "All",
        "filter_take": "✅ Take",
        "filter_stop": "❌ Stop",
        "filter_bu": "⚖️ BE",
        "filter_asset": "💰 By asset",
        "filter_date": "📅 By date",
        "filter_clear": "🗑 Clear filters",
        "filter_date_day": "Last day",
        "filter_date_week": "Last week",
        "filter_date_month": "Last month",
        "select_asset_filter": "💰 **Select asset:**",
        "select_date_filter": "📅 **Select period:**",
        "trade_detail": "📋 **Trade #{id}**\n\n🪙 Asset: {asset}\n📈 Direction: {direction}\n💰 Entry: ${entry}\n💰 Exit: ${exit}\n📊 Volume: {volume}\n💵 P&L: ${pnl}\n🎯 Outcome: {result}\n📅 Date: {date}\n😊 Emotion: {emotion}\n🔗 Links:\n{links}\n📝 Comment: {comment}",
        "recent_trades": "📋 **Trades** (page {page}/{total_pages})",
        "prev": "⬅️ Prev",
        "next": "Next ➡️",
        "refresh": "🔄 Refresh",
        "sort": "📅 Sort",
        "filter": "🔍 Filter",
        "edit": "✏️ Edit",
        "delete": "🗑 Delete",
        "back_to_list": "🔙 Back to list",
        "stats_header": "📊 Your statistics",
        "bt_period_start": "📅 Enter START date (DD.MM.YYYY):",
        "bt_period_end": "📅 Enter END date (DD.MM.YYYY):",
        "bt_timeframe": "⏱ Enter timeframe (M5, H1, H4, D1, W1):",
        "enter_exit_price_bt": "💰 Enter exit price:",
        "enter_link_bt": "🔗 Screenshot link (0 if none):",
        "error_number": "❌ Error! Enter a number.",
        "error_date": "❌ Error! Enter date in DD.MM.YYYY format.",
        "select_language": "🌐 Select language:",
        "language_set": "✅ Language set: Russian",
        "language_set_en": "✅ Language set: English",
        "support_info": "📞 **Support**\n\nContact: @your_username",
        "change_lang": "🌐 Change language",
        "support": "📞 Support",
        "settings_menu": "⚙️ **Settings**",
        "stats_menu": "📊 **Statistics Menu**\n\nSelect statistic type:",
        "stats_all": "📈 All Statistics",
        "stats_by_asset": "💰 By Asset",
        "stats_by_date": "📅 By Date",
        "stats_by_emotion": "😊 By Emotion",
        "stats_recent": "🕒 Recent Trades",
        "stats_sort": "🔄 Sort (Newest/Oldest)",
        "stats_back": "🔙 Back to Mode Menu",
        "bt_stats_header": "📊 Backtest statistics",
        "bt_total_trades": "📋 Total backtests: {total}",
        "select_asset": "💰 **Select asset:**",
        "select_period": "📅 **Select period:**",
        "period_day": "📆 Day",
        "period_week": "📅 Week",
        "period_month": "📊 Month"
    }
}

def get_text(lang, key, **kwargs):
    t = TEXTS.get(lang, TEXTS["ru"]).get(key, key)
    if kwargs:
        t = t.format(**kwargs)
    return t

# ========== БАЗЫ ДАННЫХ ==========
def init_dbs():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            asset TEXT,
            direction TEXT,
            entry_price REAL,
            exit_price REAL,
            volume REAL,
            pnl REAL,
            result TEXT,
            comment TEXT,
            trade_date TEXT,
            links TEXT,
            emotion TEXT
        )
    """)
    conn.close()
    conn = sqlite3.connect(BT_DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS backtests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            period_start TEXT,
            period_end TEXT,
            timeframe TEXT,
            commission REAL,
            spread REAL,
            asset TEXT,
            direction TEXT,
            entry_price REAL,
            exit_price REAL,
            sl_price REAL,
            tp_price REAL,
            pnl_usd REAL,
            pnl_r REAL,
            signal_quality INTEGER,
            setup TEXT,
            trigger TEXT,
            link_chart TEXT,
            entry_time TEXT,
            exit_time TEXT
        )
    """)
    conn.close()
    conn = sqlite3.connect(DB_NAME)
    conn.execute("CREATE TABLE IF NOT EXISTS users (user_id INTEGER PRIMARY KEY, lang TEXT DEFAULT 'ru')")
    conn.commit()
    conn.close()

def get_user_lang(user_id):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT lang FROM users WHERE user_id = ?", (user_id,))
    r = cur.fetchone()
    conn.close()
    return r[0] if r else "ru"

def set_user_lang(user_id, lang):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("INSERT OR REPLACE INTO users (user_id, lang) VALUES (?, ?)", (user_id, lang))
    conn.commit()
    conn.close()

def save_trade(user_id, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links, emotion):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        INSERT INTO trades (user_id, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links, emotion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links, emotion))
    conn.commit()
    conn.close()

def get_trade_by_id(trade_id, user_id):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT * FROM trades WHERE id = ? AND user_id = ?", (trade_id, user_id))
    row = cur.fetchone()
    conn.close()
    if row:
        cols = ['id', 'user_id', 'asset', 'direction', 'entry_price', 'exit_price', 'volume', 'pnl', 'result', 'comment', 'trade_date', 'links', 'emotion']
        return dict(zip(cols, row))
    return None

def update_trade_field(trade_id, user_id, field, value):
    conn = sqlite3.connect(DB_NAME)
    conn.execute(f"UPDATE trades SET {field} = ? WHERE id = ? AND user_id = ?", (value, trade_id, user_id))
    conn.commit()
    conn.close()

def delete_trade(trade_id, user_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM trades WHERE id = ? AND user_id = ?", (trade_id, user_id))
    conn.commit()
    conn.close()

def get_trades_filtered(user_id, result_filter=None, asset_filter=None, date_filter=None, sort_order="DESC"):
    conn = sqlite3.connect(DB_NAME)
    query = "SELECT * FROM trades WHERE user_id = ?"
    params = [user_id]
    if result_filter and result_filter != "all":
        if result_filter == "take":
            query += " AND result = 'TAKE'"
        elif result_filter == "stop":
            query += " AND result = 'STOP'"
        elif result_filter == "bu":
            query += " AND result = 'BU'"
    if asset_filter:
        query += " AND asset = ?"
        params.append(asset_filter)
    if date_filter:
        days = {"day": 1, "week": 7, "month": 30}
        start = (datetime.now() - timedelta(days=days[date_filter])).strftime("%Y-%m-%d")
        query += " AND trade_date >= ?"
        params.append(start)
    query += f" ORDER BY trade_date {sort_order}"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_all_assets(user_id):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT asset FROM trades WHERE user_id = ?", (user_id,))
    assets = [row[0] for row in cur.fetchall()]
    conn.close()
    return assets

def clear_trades(user_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM trades WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

def get_backtests(user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    df = pd.read_sql_query("SELECT * FROM backtests WHERE user_id = ?", conn, params=(user_id,))
    conn.close()
    return df

def clear_backtests(user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    conn.execute("DELETE FROM backtests WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

# ========== EXCEL ==========
def export_real_to_excel(df, user_id):
    if df.empty:
        return None
    df_exp = df.copy()
    df_exp = df_exp[['trade_date', 'asset', 'direction', 'entry_price', 'exit_price', 'volume', 'pnl', 'result', 'comment', 'links', 'emotion']]
    df_exp.columns = ['📅 Дата', '🪙 Актив', '📈 Направление', '💰 Вход', '💰 Выход', '📊 Объём', '💵 P&L', '🎯 Исход', '📝 Комментарий', '🔗 Ссылки', '😊 Эмоции']
    df_exp['📈 Направление'] = df_exp['📈 Направление'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_exp['🎯 Исход'] = df_exp['🎯 Исход'].replace({'TAKE': '✅ Тейк', 'STOP': '❌ Стоп', 'BU': '⚖️ БУ'})
    df_exp = df_exp.sort_values('📅 Дата', ascending=False)
    fname = f"real_{user_id}.xlsx"
    with pd.ExcelWriter(fname, engine='openpyxl') as w:
        df_exp.to_excel(w, sheet_name='Реальная торговля', index=False)
        ws = w.sheets['Реальная торговля']
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2b6cb0", end_color="2b6cb0", fill_type="solid")
        for col in range(1, len(df_exp.columns)+1):
            ws.cell(row=1, column=col).font = header_font
            ws.cell(row=1, column=col).fill = header_fill
        green = PatternFill(start_color="c6f7d0", end_color="c6f7d0", fill_type="solid")
        red = PatternFill(start_color="fecaca", end_color="fecaca", fill_type="solid")
        yellow = PatternFill(start_color="fff3cd", end_color="fff3cd", fill_type="solid")
        for row in range(2, len(df_exp)+2):
            val = ws.cell(row=row, column=7).value
            if val and val > 0:
                for col in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=col).fill = green
            elif val and val < 0:
                for col in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=col).fill = red
            elif val == 0:
                for col in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=col).fill = yellow
        for col in range(1, len(df_exp.columns)+1):
            max_len = 0
            col_letter = get_column_letter(col)
            for row in range(1, len(df_exp)+2):
                v = ws.cell(row=row, column=col).value
                if v:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[col_letter].width = min(max_len+2, 30)
        ws.freeze_panes = 'A2'
    return fname

def export_backtest_to_excel(df, user_id):
    if df.empty:
        return None
    df_exp = df.copy()
    df_exp = df_exp[['period_start', 'period_end', 'timeframe', 'asset', 'direction', 'entry_price', 'exit_price', 'pnl_usd', 'pnl_r', 'signal_quality', 'setup', 'trigger']]
    df_exp.columns = ['📅 Начало', '📅 Конец', '⏱ Таймфрейм', '🪙 Актив', '📈 Направление', '💰 Вход', '💰 Выход', '💵 P&L', '📊 P&L (R)', '⭐ Качество', '🎯 Сетап', '⚡ Триггер']
    df_exp['📈 Направление'] = df_exp['📈 Направление'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_exp = df_exp.sort_values('📅 Начало', ascending=False)
    fname = f"backtest_{user_id}.xlsx"
    with pd.ExcelWriter(fname, engine='openpyxl') as w:
        df_exp.to_excel(w, sheet_name='Бэктест', index=False)
        ws = w.sheets['Бэктест']
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2b6cb0", end_color="2b6cb0", fill_type="solid")
        for col in range(1, len(df_exp.columns)+1):
            ws.cell(row=1, column=col).font = header_font
            ws.cell(row=1, column=col).fill = header_fill
        green = PatternFill(start_color="c6f7d0", end_color="c6f7d0", fill_type="solid")
        red = PatternFill(start_color="fecaca", end_color="fecaca", fill_type="solid")
        for row in range(2, len(df_exp)+2):
            val = ws.cell(row=row, column=8).value
            if val and val > 0:
                for col in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=col).fill = green
            elif val and val < 0:
                for col in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=col).fill = red
        for col in range(1, len(df_exp.columns)+1):
            max_len = 0
            col_letter = get_column_letter(col)
            for row in range(1, len(df_exp)+2):
                v = ws.cell(row=row, column=col).value
                if v:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[col_letter].width = min(max_len+2, 30)
        ws.freeze_panes = 'A2'
    return fname

# ========== СТАТИСТИКА ==========
def get_real_stats_text(df, lang, title_key="stats_header"):
    if df.empty:
        return get_text(lang, "no_data")
    total = len(df)
    wins = len(df[df['pnl'] > 0])
    losses = len(df[df['pnl'] < 0])
    bu = len(df[df['pnl'] == 0])
    wr = wins/total*100 if total else 0
    longs = len(df[df['direction'] == 'LONG'])
    shorts = len(df[df['direction'] == 'SHORT'])
    total_pnl = df['pnl'].sum()
    avg_pnl = df['pnl'].mean()
    best = df['pnl'].max()
    worst = df['pnl'].min()
    sum_profit = df[df['pnl']>0]['pnl'].sum()
    sum_loss = abs(df[df['pnl']<0]['pnl'].sum())
    pf = sum_profit/sum_loss if sum_loss else sum_profit
    emotions = df['emotion'].value_counts().to_dict()
    emotion_text = "\n".join([f"{e}: {c}" for e, c in emotions.items()]) if emotions else get_text(lang, "no_data")
    return (
        f"{get_text(lang, title_key)}\n\n"
        f"📋 Всего сделок: {total}\n"
        f"✅ Тейков: {wins}\n"
        f"❌ Стопов: {losses}\n"
        f"⚖️ БУ: {bu}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"📈 Лонги: {longs} | 📉 Шорты: {shorts}\n"
        f"💰 Суммарный P&L: ${total_pnl:.2f}\n"
        f"📊 Средняя сделка: ${avg_pnl:.2f}\n"
        f"🏆 Лучшая: +${best:.2f}\n"
        f"💀 Худшая: ${worst:.2f}\n"
        f"⚙️ Профит-фактор: {pf:.2f}\n\n"
        f"😊 **Эмоции:**\n{emotion_text}"
    )

def get_backtest_stats_text(df, lang):
    if df.empty:
        return get_text(lang, "no_data")
    total = len(df)
    wins = len(df[df['pnl_usd'] > 0])
    losses = len(df[df['pnl_usd'] < 0])
    wr = wins/total*100 if total else 0
    avg_r = df['pnl_r'].mean()
    total_r = df['pnl_r'].sum()
    avg_q = df['signal_quality'].mean()
    return (
        f"{get_text(lang, 'bt_stats_header')}\n\n"
        f"{get_text(lang, 'bt_total_trades', total=total)}\n"
        f"✅ Прибыльных: {wins}\n"
        f"❌ Убыточных: {losses}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"📊 Средний R: {avg_r:.2f}\n"
        f"💰 Суммарный R: {total_r:.2f}\n"
        f"⭐ Качество сигнала: {avg_q:.1f}/5"
    )

# ========== КЛАВИАТУРЫ ==========
def mode_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "mode_real"), callback_data="mode_real")],
        [InlineKeyboardButton(text=get_text(lang, "mode_backtest"), callback_data="mode_backtest")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu")]
    ])

def real_menu_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "add_trade"), callback_data="real_add_trade")],
        [InlineKeyboardButton(text=get_text(lang, "list_trades"), callback_data="real_list_trades")],
        [InlineKeyboardButton(text=get_text(lang, "stats"), callback_data="stats_menu")],
        [InlineKeyboardButton(text=get_text(lang, "excel_real"), callback_data="get_real_excel")],
        [InlineKeyboardButton(text=get_text(lang, "clear"), callback_data="clear_confirm")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def backtest_menu_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "add_trade"), callback_data="backtest_add_trade")],
        [InlineKeyboardButton(text=get_text(lang, "list_trades"), callback_data="backtest_list_trades")],
        [InlineKeyboardButton(text=get_text(lang, "stats"), callback_data="stats_menu")],
        [InlineKeyboardButton(text=get_text(lang, "excel_backtest"), callback_data="get_backtest_excel")],
        [InlineKeyboardButton(text=get_text(lang, "clear"), callback_data="clear_confirm")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def direction_kb(lang, prefix):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "long"), callback_data=f"{prefix}_LONG"),
         InlineKeyboardButton(text=get_text(lang, "short"), callback_data=f"{prefix}_SHORT")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def result_kb(lang, prefix):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "take"), callback_data=f"{prefix}_TAKE"),
         InlineKeyboardButton(text=get_text(lang, "stop"), callback_data=f"{prefix}_STOP"),
         InlineKeyboardButton(text=get_text(lang, "bu"), callback_data=f"{prefix}_BU")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def emotion_kb(lang, trade_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "emotion_calm"), callback_data=f"emotion_{trade_id}_calm")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_fear"), callback_data=f"emotion_{trade_id}_fear")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_greed"), callback_data=f"emotion_{trade_id}_greed")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_tilt"), callback_data=f"emotion_{trade_id}_tilt")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_confidence"), callback_data=f"emotion_{trade_id}_confidence")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def link_yesno_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "yes"), callback_data="link_yes")],
        [InlineKeyboardButton(text=get_text(lang, "no"), callback_data="link_no")]
    ])

def back_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def confirm_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "confirm_clear"), callback_data="clear_yes")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def settings_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "change_lang"), callback_data="change_lang")],
        [InlineKeyboardButton(text=get_text(lang, "support"), callback_data="support")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def lang_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru")],
        [InlineKeyboardButton(text="🇬🇧 English", callback_data="lang_en")]
    ])

def trades_list_kb(trades, page, total_pages, lang, sort_order, result_filter, asset_filter, date_filter):
    buttons = []
    for _, row in trades.iterrows():
        pnl = row['pnl']
        if pnl > 0:
            emoji = "✅"
            pnl_text = f"+${pnl:.0f}"
        elif pnl < 0:
            emoji = "❌"
            pnl_text = f"${pnl:.0f}"
        else:
            emoji = "⚖️"
            pnl_text = "БУ"
        if row['result'] == "TAKE":
            res_emoji = "✅"
        elif row['result'] == "STOP":
            res_emoji = "❌"
        else:
            res_emoji = "⚖️"
        buttons.append([InlineKeyboardButton(text=f"{row['asset']} {res_emoji} {pnl_text}", callback_data=f"trade_{row['id']}")])
    nav = []
    if page > 1:
        nav.append(InlineKeyboardButton(text=get_text(lang, "prev"), callback_data=f"trades_page_{page-1}"))
    if page < total_pages:
        nav.append(InlineKeyboardButton(text=get_text(lang, "next"), callback_data=f"trades_page_{page+1}"))
    sort_text = get_text(lang, "sort_newest") if sort_order == "DESC" else get_text(lang, "sort_oldest")
    filter_btns = [
        InlineKeyboardButton(text=f"{get_text(lang, 'sort')}: {sort_text}", callback_data="sort_menu"),
        InlineKeyboardButton(text=get_text(lang, "filter"), callback_data="filter_menu")
    ]
    if nav:
        buttons.append(nav)
    buttons.append(filter_btns)
    buttons.append([InlineKeyboardButton(text=get_text(lang, "refresh"), callback_data="real_list_trades"),
                   InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def sort_menu_kb(lang, current_sort):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"{'✅ ' if current_sort == 'DESC' else ''}{get_text(lang, 'sort_newest')}", callback_data="sort_newest")],
        [InlineKeyboardButton(text=f"{'✅ ' if current_sort == 'ASC' else ''}{get_text(lang, 'sort_oldest')}", callback_data="sort_oldest")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="list_trades")]
    ])

def filter_menu_kb(lang, result_filter, asset_filter, date_filter, assets):
    buttons = [
        [InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'all' else ''}{get_text(lang, 'filter_all')}", callback_data="filter_result_all")],
        [InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'take' else ''}{get_text(lang, 'filter_take')}", callback_data="filter_result_take")],
        [InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'stop' else ''}{get_text(lang, 'filter_stop')}", callback_data="filter_result_stop")],
        [InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'bu' else ''}{get_text(lang, 'filter_bu')}", callback_data="filter_result_bu")],
        [InlineKeyboardButton(text=get_text(lang, "filter_asset"), callback_data="filter_asset_menu")],
        [InlineKeyboardButton(text=get_text(lang, "filter_date"), callback_data="filter_date_menu")],
        [InlineKeyboardButton(text=get_text(lang, "filter_clear"), callback_data="filter_clear")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="list_trades")]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def asset_filter_kb(assets, lang):
    btns = [[InlineKeyboardButton(text=a, callback_data=f"filter_asset_{a}")] for a in assets]
    btns.append([InlineKeyboardButton(text=get_text(lang, "back"), callback_data="filter_menu")])
    return InlineKeyboardMarkup(inline_keyboard=btns)

def date_filter_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "filter_date_day"), callback_data="filter_date_day")],
        [InlineKeyboardButton(text=get_text(lang, "filter_date_week"), callback_data="filter_date_week")],
        [InlineKeyboardButton(text=get_text(lang, "filter_date_month"), callback_data="filter_date_month")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="filter_menu")]
    ])

def trade_detail_kb(trade_id, lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "edit"), callback_data=f"edit_{trade_id}"),
         InlineKeyboardButton(text=get_text(lang, "delete"), callback_data=f"delete_{trade_id}")],
        [InlineKeyboardButton(text=get_text(lang, "back_to_list"), callback_data="real_list_trades")]
    ])

def edit_field_kb(trade_id, lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "edit_asset"), callback_data=f"edit_field_{trade_id}_asset")],
        [InlineKeyboardButton(text=get_text(lang, "edit_direction"), callback_data=f"edit_field_{trade_id}_direction")],
        [InlineKeyboardButton(text=get_text(lang, "edit_entry_price"), callback_data=f"edit_field_{trade_id}_entry_price")],
        [InlineKeyboardButton(text=get_text(lang, "edit_exit_price"), callback_data=f"edit_field_{trade_id}_exit_price")],
        [InlineKeyboardButton(text=get_text(lang, "edit_volume"), callback_data=f"edit_field_{trade_id}_volume")],
        [InlineKeyboardButton(text=get_text(lang, "edit_result"), callback_data=f"edit_field_{trade_id}_result")],
        [InlineKeyboardButton(text=get_text(lang, "edit_comment"), callback_data=f"edit_field_{trade_id}_comment")],
        [InlineKeyboardButton(text=get_text(lang, "edit_date"), callback_data=f"edit_field_{trade_id}_trade_date")],
        [InlineKeyboardButton(text=get_text(lang, "edit_emotion"), callback_data=f"edit_field_{trade_id}_emotion")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data=f"trade_{trade_id}")]
    ])

def edit_direction_kb(trade_id, lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "long"), callback_data=f"edit_value_{trade_id}_direction_LONG"),
         InlineKeyboardButton(text=get_text(lang, "short"), callback_data=f"edit_value_{trade_id}_direction_SHORT")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data=f"edit_{trade_id}")]
    ])

def edit_result_kb(trade_id, lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "take"), callback_data=f"edit_value_{trade_id}_result_TAKE"),
         InlineKeyboardButton(text=get_text(lang, "stop"), callback_data=f"edit_value_{trade_id}_result_STOP"),
         InlineKeyboardButton(text=get_text(lang, "bu"), callback_data=f"edit_value_{trade_id}_result_BU")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data=f"edit_{trade_id}")]
    ])

def edit_emotion_kb(trade_id, lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "emotion_calm"), callback_data=f"edit_value_{trade_id}_emotion_😊 Спокойствие")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_fear"), callback_data=f"edit_value_{trade_id}_emotion_😨 Страх")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_greed"), callback_data=f"edit_value_{trade_id}_emotion_😈 Жадность")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_tilt"), callback_data=f"edit_value_{trade_id}_emotion_🤬 Тильт")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_confidence"), callback_data=f"edit_value_{trade_id}_emotion_😌 Уверенность")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data=f"edit_{trade_id}")]
    ])

def stats_menu_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "stats_all"), callback_data="stats_all")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_asset"), callback_data="stats_by_asset")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_date"), callback_data="stats_by_date")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_emotion"), callback_data="stats_by_emotion")],
        [InlineKeyboardButton(text=get_text(lang, "stats_recent"), callback_data="stats_recent")],
        [InlineKeyboardButton(text=get_text(lang, "stats_sort"), callback_data="stats_sort")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

# ========== FSM ==========
class RealTradeForm(StatesGroup):
    asset = State()
    direction = State()
    entry_price = State()
    exit_price = State()
    volume = State()
    result = State()
    comment = State()
    add_link = State()
    link_tf = State()
    link_url = State()
    links = State()
    trade_date = State()
    emotion = State()

class BacktestForm(StatesGroup):
    period_start = State()
    period_end = State()
    timeframe = State()
    commission = State()
    spread = State()
    asset = State()
    direction = State()
    entry_price = State()
    sl_price = State()
    tp_price = State()
    exit_price = State()
    entry_time = State()
    exit_time = State()
    setup = State()
    trigger = State()
    signal_quality = State()
    link_chart = State()

class EditForm(StatesGroup):
    waiting_for_value = State()

# ========== ОБРАБОТЧИКИ ==========
bot = None
dp = Dispatcher()

# ---------- СТАРТ ----------
@dp.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await state.clear()
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    if not lang:
        await msg.answer(get_text("ru", "select_language"), reply_markup=lang_kb())
        return
    await msg.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))

@dp.callback_query(F.data.startswith("lang_"))
async def set_lang(call: CallbackQuery, state: FSMContext):
    lang = call.data.split("_")[1]
    set_user_lang(call.from_user.id, lang)
    await call.message.delete()
    await call.message.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "mode_real")
async def mode_real(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=real_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "mode_backtest")
async def mode_backtest(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=backtest_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "back_mode")
async def back_mode(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

# ---------- НАСТРОЙКИ ----------
@dp.callback_query(F.data == "settings_menu")
async def settings_menu(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "settings_menu"), parse_mode="Markdown", reply_markup=settings_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "change_lang")
async def change_lang(call: CallbackQuery):
    await call.message.edit_text(get_text("ru", "select_language"), reply_markup=lang_kb())
    await call.answer()

@dp.callback_query(F.data == "support")
async def support(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "support_info"), parse_mode="Markdown", reply_markup=settings_kb(lang))
    await call.answer()

# ---------- РЕАЛЬНАЯ ТОРГОВЛЯ ----------
@dp.callback_query(F.data == "real_add_trade")
async def real_add_trade(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(RealTradeForm.asset)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(RealTradeForm.asset)
async def real_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(RealTradeForm.direction)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang, "real"))

@dp.callback_query(F.data.startswith("real_LONG") | F.data.startswith("real_SHORT"))
async def real_direction(call: CallbackQuery, state: FSMContext):
    direction = call.data.split("_")[1]
    await state.update_data(direction=direction)
    await state.set_state(RealTradeForm.entry_price)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(RealTradeForm.entry_price)
async def real_entry(msg: Message, state: FSMContext):
    try:
        value = msg.text.replace(",", ".")
        await state.update_data(entry_price=float(value))
        await state.set_state(RealTradeForm.exit_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_exit_price"), reply_markup=back_kb(lang))
    except ValueError:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"), reply_markup=back_kb(lang))

@dp.message(RealTradeForm.exit_price)
async def real_exit(msg: Message, state: FSMContext):
    try:
        value = msg.text.replace(",", ".")
        await state.update_data(exit_price=float(value))
        await state.set_state(RealTradeForm.volume)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_volume"), reply_markup=back_kb(lang))
    except ValueError:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"), reply_markup=back_kb(lang))

@dp.message(RealTradeForm.volume)
async def real_volume(msg: Message, state: FSMContext):
    try:
        value = msg.text.replace(",", ".")
        await state.update_data(volume=float(value))
        await state.set_state(RealTradeForm.result)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "choose_result"), reply_markup=result_kb(lang, "real"))
    except ValueError:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"), reply_markup=back_kb(lang))

@dp.callback_query(F.data.startswith("real_TAKE") | F.data.startswith("real_STOP") | F.data.startswith("real_BU"))
async def real_result(call: CallbackQuery, state: FSMContext):
    result = call.data.split("_")[1]
    await state.update_data(result=result)
    await state.set_state(RealTradeForm.comment)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_comment"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(RealTradeForm.comment)
async def real_comment(msg: Message, state: FSMContext):
    com = msg.text.strip()
    if com == "-":
        com = ""
    await state.update_data(comment=com)
    await state.set_state(RealTradeForm.add_link)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "add_link_question"), reply_markup=link_yesno_kb(lang))

@dp.callback_query(F.data == "link_yes")
async def link_yes(call: CallbackQuery, state: FSMContext):
    await state.set_state(RealTradeForm.link_url)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_link"), reply_markup=back_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "link_no")
async def link_no(call: CallbackQuery, state: FSMContext):
    await state.update_data(links="")
    await state.set_state(RealTradeForm.trade_date)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_date"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(RealTradeForm.link_url)
async def real_get_link(msg: Message, state: FSMContext):
    await state.update_data(link_url=msg.text)
    await state.set_state(RealTradeForm.link_tf)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_timeframe"), reply_markup=back_kb(lang))

@dp.message(RealTradeForm.link_tf)
async def real_get_tf(msg: Message, state: FSMContext):
    tf = msg.text
    data = await state.get_data()
    links = data.get("links", "")
    new_link = f"{tf}: {data.get('link_url')}"
    links = f"{links}\n{new_link}" if links else new_link
    await state.update_data(links=links)
    await state.set_state(RealTradeForm.add_link)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "link_saved"), reply_markup=link_yesno_kb(lang))

@dp.message(RealTradeForm.trade_date)
async def real_date(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id)
    dstr = msg.text.strip().lower()
    if dstr in ["сегодня", "today"]:
        trade_date = datetime.now().strftime("%Y-%m-%d")
    else:
        try:
            trade_date = datetime.strptime(dstr, "%d.%m.%Y").strftime("%Y-%m-%d")
        except:
            await msg.answer(get_text(lang, "error_date"))
            return
    await state.update_data(trade_date=trade_date)
    await state.set_state(RealTradeForm.emotion)
    await msg.answer(get_text(lang, "enter_emotion"), reply_markup=emotion_kb(lang, "temp"))

@dp.callback_query(F.data.startswith("emotion_temp_"))
async def real_emotion(call: CallbackQuery, state: FSMContext):
    emotion_map = {
        "calm": "😊 Спокойствие",
        "fear": "😨 Страх",
        "greed": "😈 Жадность",
        "tilt": "🤬 Тильт",
        "confidence": "😌 Уверенность"
    }
    emotion_key = call.data.split("_")[2]
    emotion = emotion_map.get(emotion_key, "😊 Спокойствие")
    data = await state.get_data()
    direction = data['direction']
    entry = data['entry_price']
    exit_p = data['exit_price']
    vol = data['volume']
    pnl = (exit_p - entry) * vol if direction == "LONG" else (entry - exit_p) * vol
    if data['result'] == "BU":
        pnl = 0
    save_trade(
        user_id=call.from_user.id,
        asset=data['asset'], direction=direction,
        entry_price=entry, exit_price=exit_p, volume=vol, pnl=pnl,
        result=data['result'], comment=data['comment'], trade_date=data['trade_date'],
        links=data.get('links', ''), emotion=emotion
    )
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "trade_saved"), parse_mode="Markdown", reply_markup=real_menu_kb(lang))
    await call.answer()

# ---------- СПИСОК СДЕЛОК ----------
@dp.callback_query(F.data == "real_list_trades")
async def real_list_trades(call: CallbackQuery, state: FSMContext):
    await state.clear()
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    page = data.get('page', 1)
    sort_order = data.get('sort_order', 'DESC')
    result_filter = data.get('result_filter', 'all')
    asset_filter = data.get('asset_filter', None)
    date_filter = data.get('date_filter', None)
    df = get_trades_filtered(uid, result_filter, asset_filter, date_filter, sort_order)
    if df.empty:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    total = len(df)
    pages = (total + TRADES_PER_PAGE - 1) // TRADES_PER_PAGE
    if page > pages:
        page = pages
    start = (page - 1) * TRADES_PER_PAGE
    end = start + TRADES_PER_PAGE
    trades_df = df.iloc[start:end]
    await state.update_data(page=page, sort_order=sort_order, result_filter=result_filter, asset_filter=asset_filter, date_filter=date_filter)
    text = get_text(lang, "recent_trades", page=page, total_pages=pages)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=trades_list_kb(trades_df, page, pages, lang, sort_order, result_filter, asset_filter, date_filter))
    await call.answer()

@dp.callback_query(F.data.startswith("trades_page_"))
async def trades_page(call: CallbackQuery, state: FSMContext):
    page = int(call.data.split("_")[2])
    await state.update_data(page=page)
    await real_list_trades(call, state)

@dp.callback_query(F.data == "sort_menu")
async def sort_menu(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    lang = get_user_lang(call.from_user.id)
    sort_order = data.get('sort_order', 'DESC')
    await call.message.edit_text(get_text(lang, "sort_label"), parse_mode="Markdown", reply_markup=sort_menu_kb(lang, sort_order))
    await call.answer()

@dp.callback_query(F.data == "sort_newest")
async def sort_newest(call: CallbackQuery, state: FSMContext):
    await state.update_data(sort_order="DESC", page=1)
    await real_list_trades(call, state)

@dp.callback_query(F.data == "sort_oldest")
async def sort_oldest(call: CallbackQuery, state: FSMContext):
    await state.update_data(sort_order="ASC", page=1)
    await real_list_trades(call, state)

@dp.callback_query(F.data == "filter_menu")
async def filter_menu(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    uid = call.from_user.id
    lang = get_user_lang(uid)
    result_filter = data.get('result_filter', 'all')
    asset_filter = data.get('asset_filter', None)
    date_filter = data.get('date_filter', None)
    assets = get_all_assets(uid)
    await call.message.edit_text(get_text(lang, "filter_label"), parse_mode="Markdown", reply_markup=filter_menu_kb(lang, result_filter, asset_filter, date_filter, assets))
    await call.answer()

@dp.callback_query(F.data.startswith("filter_result_"))
async def filter_result(call: CallbackQuery, state: FSMContext):
    result = call.data.split("_")[2]
    await state.update_data(result_filter=result, page=1)
    await real_list_trades(call, state)

@dp.callback_query(F.data == "filter_asset_menu")
async def filter_asset_menu(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    assets = get_all_assets(uid)
    if not assets:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    await call.message.edit_text(get_text(lang, "select_asset_filter"), parse_mode="Markdown", reply_markup=asset_filter_kb(assets, lang))
    await call.answer()

@dp.callback_query(F.data.startswith("filter_asset_"))
async def filter_asset_apply(call: CallbackQuery, state: FSMContext):
    asset = call.data.split("_")[2]
    await state.update_data(asset_filter=asset, page=1)
    await real_list_trades(call, state)

@dp.callback_query(F.data == "filter_date_menu")
async def filter_date_menu(call: CallbackQuery, state: FSMContext):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_date_filter"), parse_mode="Markdown", reply_markup=date_filter_kb(lang))
    await call.answer()

@dp.callback_query(F.data.startswith("filter_date_"))
async def filter_date_apply(call: CallbackQuery, state: FSMContext):
    date_filter = call.data.split("_")[2]
    await state.update_data(date_filter=date_filter, page=1)
    await real_list_trades(call, state)

@dp.callback_query(F.data == "filter_clear")
async def filter_clear(call: CallbackQuery, state: FSMContext):
    await state.update_data(result_filter="all", asset_filter=None, date_filter=None, page=1)
    await real_list_trades(call, state)

@dp.callback_query(F.data.startswith("trade_"))
async def show_trade_detail(call: CallbackQuery, state: FSMContext):
    trade_id = int(call.data.split("_")[1])
    uid = call.from_user.id
    lang = get_user_lang(uid)
    trade = get_trade_by_id(trade_id, uid)
    if not trade:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    links = trade.get('links', '') or '-'
    result_map = {"TAKE": "✅ Тейк", "STOP": "❌ Стоп", "BU": "⚖️ БУ"}
    dir_map = {"LONG": "🟢 LONG", "SHORT": "🔴 SHORT"}
    text = get_text(lang, "trade_detail",
        id=trade['id'], asset=trade['asset'], direction=dir_map.get(trade['direction'], trade['direction']),
        entry=trade['entry_price'], exit=trade['exit_price'], volume=trade['volume'],
        pnl=trade['pnl'], result=result_map.get(trade['result'], trade['result']),
        date=trade['trade_date'], emotion=trade['emotion'], links=links, comment=trade['comment'] or "-"
    )
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=trade_detail_kb(trade_id, lang))
    await call.answer()

# ---------- УДАЛЕНИЕ ----------
@dp.callback_query(F.data.startswith("delete_"))
async def delete_confirm(call: CallbackQuery, state: FSMContext):
    trade_id = int(call.data.split("_")[1])
    lang = get_user_lang(call.from_user.id)
    await state.update_data(delete_id=trade_id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "yes"), callback_data="delete_yes"),
         InlineKeyboardButton(text=get_text(lang, "no"), callback_data="real_list_trades")]
    ])
    await call.message.edit_text(get_text(lang, "confirm_delete", id=trade_id), parse_mode="Markdown", reply_markup=kb)
    await call.answer()

@dp.callback_query(F.data == "delete_yes")
async def delete_execute(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    trade_id = data.get('delete_id')
    if trade_id:
        delete_trade(trade_id, call.from_user.id)
    await state.update_data(delete_id=None)
    await real_list_trades(call, state)

# ---------- РЕДАКТИРОВАНИЕ ----------
@dp.callback_query(F.data.startswith("edit_") and not F.data.startswith("edit_field_") and not F.data.startswith("edit_value_"))
async def edit_menu(call: CallbackQuery, state: FSMContext):
    trade_id = int(call.data.split("_")[1])
    await state.update_data(edit_id=trade_id)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "edit_field_select", id=trade_id), parse_mode="Markdown", reply_markup=edit_field_kb(trade_id, lang))
    await call.answer()

@dp.callback_query(F.data.startswith("edit_field_"))
async def edit_field(call: CallbackQuery, state: FSMContext):
    parts = call.data.split("_")
    trade_id = int(parts[2])
    field = parts[3]
    lang = get_user_lang(call.from_user.id)
    await state.update_data(edit_id=trade_id, edit_field=field)
    if field == "direction":
        await call.message.edit_text(get_text(lang, "choose_direction"), parse_mode="Markdown", reply_markup=edit_direction_kb(trade_id, lang))
    elif field == "result":
        await call.message.edit_text(get_text(lang, "choose_result"), parse_mode="Markdown", reply_markup=edit_result_kb(trade_id, lang))
    elif field == "emotion":
        await call.message.edit_text(get_text(lang, "enter_emotion"), parse_mode="Markdown", reply_markup=edit_emotion_kb(trade_id, lang))
    else:
        field_names = {"asset": get_text(lang, "edit_asset"), "entry_price": get_text(lang, "edit_entry_price"),
                       "exit_price": get_text(lang, "edit_exit_price"), "volume": get_text(lang, "edit_volume"),
                       "comment": get_text(lang, "edit_comment"), "trade_date": get_text(lang, "edit_date")}
        await state.set_state(EditForm.waiting_for_value)
        await call.message.edit_text(get_text(lang, "enter_new_value", field=field_names.get(field, field)), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data=f"edit_{trade_id}")]]))
    await call.answer()

@dp.callback_query(F.data.startswith("edit_value_"))
async def edit_value(call: CallbackQuery, state: FSMContext):
    parts = call.data.split("_")
    trade_id = int(parts[2])
    field = parts[3]
    value = "_".join(parts[4:])
    lang = get_user_lang(call.from_user.id)
    if field == "direction":
        update_trade_field(trade_id, call.from_user.id, "direction", value)
        await call.message.edit_text(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back_to_list"), callback_data="real_list_trades")]]))
    elif field == "result":
        update_trade_field(trade_id, call.from_user.id, "result", value)
        if value == "BU":
            update_trade_field(trade_id, call.from_user.id, "pnl", 0)
        await call.message.edit_text(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back_to_list"), callback_data="real_list_trades")]]))
    elif field == "emotion":
        update_trade_field(trade_id, call.from_user.id, "emotion", value)
        await call.message.edit_text(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back_to_list"), callback_data="real_list_trades")]]))
    await call.answer()

@dp.message(EditForm.waiting_for_value)
async def edit_text_value(msg: Message, state: FSMContext):
    data = await state.get_data()
    trade_id = data.get('edit_id')
    field = data.get('edit_field')
    value = msg.text.strip()
    lang = get_user_lang(msg.from_user.id)
    
    if not trade_id or not field:
        await msg.answer("❌ Ошибка: не найдена сделка. Попробуйте снова.")
        await state.clear()
        return
    
    if not value:
        await msg.answer("❌ Ошибка: значение не может быть пустым.")
        return
    
    if field == "asset":
        update_trade_field(trade_id, msg.from_user.id, "asset", value.upper())
        await msg.answer(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown")
    elif field == "entry_price":
        try:
            val = float(value.replace(",", "."))
            update_trade_field(trade_id, msg.from_user.id, "entry_price", val)
            trade = get_trade_by_id(trade_id, msg.from_user.id)
            if trade and trade['result'] != "BU":
                if trade['direction'] == "LONG":
                    new_pnl = (trade['exit_price'] - val) * trade['volume']
                else:
                    new_pnl = (val - trade['exit_price']) * trade['volume']
                update_trade_field(trade_id, msg.from_user.id, "pnl", new_pnl)
            await msg.answer(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown")
        except ValueError:
            await msg.answer(get_text(lang, "error_number"))
            return
    elif field == "exit_price":
        try:
            val = float(value.replace(",", "."))
            update_trade_field(trade_id, msg.from_user.id, "exit_price", val)
            trade = get_trade_by_id(trade_id, msg.from_user.id)
            if trade and trade['result'] != "BU":
                if trade['direction'] == "LONG":
                    new_pnl = (val - trade['entry_price']) * trade['volume']
                else:
                    new_pnl = (trade['entry_price'] - val) * trade['volume']
                update_trade_field(trade_id, msg.from_user.id, "pnl", new_pnl)
            await msg.answer(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown")
        except ValueError:
            await msg.answer(get_text(lang, "error_number"))
            return
    elif field == "volume":
        try:
            val = float(value.replace(",", "."))
            update_trade_field(trade_id, msg.from_user.id, "volume", val)
            trade = get_trade_by_id(trade_id, msg.from_user.id)
            if trade and trade['result'] != "BU":
                if trade['direction'] == "LONG":
                    new_pnl = (trade['exit_price'] - trade['entry_price']) * val
                else:
                    new_pnl = (trade['entry_price'] - trade['exit_price']) * val
                update_trade_field(trade_id, msg.from_user.id, "pnl", new_pnl)
            await msg.answer(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown")
        except ValueError:
            await msg.answer(get_text(lang, "error_number"))
            return
    elif field == "comment":
        update_trade_field(trade_id, msg.from_user.id, "comment", value)
        await msg.answer(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown")
    elif field == "trade_date":
        try:
            new_date = datetime.strptime(value, "%d.%m.%Y").strftime("%Y-%m-%d")
            update_trade_field(trade_id, msg.from_user.id, "trade_date", new_date)
            await msg.answer(get_text(lang, "field_updated", field=get_text(lang, f"edit_{field}")), parse_mode="Markdown")
        except ValueError:
            await msg.answer(get_text(lang, "error_date"))
            return
    
    await state.clear()
    await msg.answer("✅ Редактирование завершено", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back_to_list"), callback_data="real_list_trades")]]))

# ---------- СТАТИСТИКА МЕНЮ ----------
@dp.callback_query(F.data == "stats_menu")
async def stats_menu(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "stats_menu"), parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "stats_all")
async def stats_all(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid)
    text = get_real_stats_text(df, lang, "stats_header")
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "stats_by_asset")
async def stats_by_asset(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    assets = get_all_assets(uid)
    if not assets:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    buttons = [[InlineKeyboardButton(text=a, callback_data=f"stats_asset_{a}")] for a in assets]
    buttons.append([InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")])
    await call.message.edit_text(get_text(lang, "select_asset"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@dp.callback_query(F.data.startswith("stats_asset_"))
async def show_asset_stats(call: CallbackQuery):
    asset = call.data.split("_")[2]
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, asset_filter=asset)
    text = get_real_stats_text(df, lang, f"📊 Статистика по {asset}")
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "stats_by_date")
async def stats_by_date(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_period"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "period_day"), callback_data="stats_date_day")],
        [InlineKeyboardButton(text=get_text(lang, "period_week"), callback_data="stats_date_week")],
        [InlineKeyboardButton(text=get_text(lang, "period_month"), callback_data="stats_date_month")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ]))
    await call.answer()

@dp.callback_query(F.data.startswith("stats_date_"))
async def show_date_stats(call: CallbackQuery):
    date_filter = call.data.split("_")[2]
    uid = call.from_user.id
    lang = get_user_lang(uid)
    days = {"day": 1, "week": 7, "month": 30}
    start = (datetime.now() - timedelta(days=days[date_filter])).strftime("%Y-%m-%d")
    df = get_trades_filtered(uid, date_filter=date_filter)
    titles = {"day": "За сегодня", "week": "За неделю", "month": "За месяц"}
    text = get_real_stats_text(df, lang, titles.get(date_filter, "Статистика"))
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "stats_by_emotion")
async def stats_by_emotion(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    emotions = ['😊 Спокойствие', '😨 Страх', '😈 Жадность', '🤬 Тильт', '😌 Уверенность']
    df = get_trades_filtered(uid)
    text = "😊 **Статистика по эмоциям:**\n\n"
    for e in emotions:
        sub = df[df['emotion'] == e]
        if not sub.empty:
            total = len(sub)
            wins = len(sub[sub['pnl'] > 0])
            wr = wins/total*100 if total else 0
            avg = sub['pnl'].mean()
            text += f"{e}: {total} сделок, винрейт {wr:.0f}%, средний ${avg:.0f}\n"
        else:
            text += f"{e}: 0 сделок\n"
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "stats_recent")
async def stats_recent(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, sort_order="DESC")
    if df.empty:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    df = df.head(10)
    text = get_text(lang, "recent_trades", count=len(df)) + "\n\n"
    for _, row in df.iterrows():
        pnl = row['pnl']
        emoji = "✅" if pnl > 0 else ("❌" if pnl < 0 else "⚖️")
        pnl_text = f"+${pnl:.0f}" if pnl > 0 else (f"${pnl:.0f}" if pnl < 0 else "БУ")
        text += f"{emoji} #{row['id']} {row['asset']} | {row['direction']} | {pnl_text} | {row['trade_date']}\n"
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "stats_sort")
async def stats_sort(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "stats_sort"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "sort_newest"), callback_data="sort_newest_stats")],
        [InlineKeyboardButton(text=get_text(lang, "sort_oldest"), callback_data="sort_oldest_stats")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ]))
    await call.answer()

@dp.callback_query(F.data == "sort_newest_stats")
async def sort_newest_stats(call: CallbackQuery, state: FSMContext):
    await state.update_data(sort_order="DESC")
    await stats_all(call)

@dp.callback_query(F.data == "sort_oldest_stats")
async def sort_oldest_stats(call: CallbackQuery, state: FSMContext):
    await state.update_data(sort_order="ASC")
    await stats_all(call)

# ---------- EXCEL ----------
@dp.callback_query(F.data == "get_real_excel")
async def get_real_excel(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid)
    if df.empty:
        await call.answer(get_text(lang, "no_data_add_trade"), show_alert=True)
        return
    fname = export_real_to_excel(df, uid)
    try:
        await call.message.answer_document(document=FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    finally:
        if os.path.exists(fname):
            os.remove(fname)
    await call.answer()

@dp.callback_query(F.data == "get_backtest_excel")
async def get_backtest_excel(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_backtests(uid)
    if df.empty:
        await call.answer(get_text(lang, "no_data_add_trade"), show_alert=True)
        return
    fname = export_backtest_to_excel(df, uid)
    try:
        await call.message.answer_document(document=FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    finally:
        if os.path.exists(fname):
            os.remove(fname)
    await call.answer()

# ---------- ОЧИСТКА ----------
@dp.callback_query(F.data == "clear_confirm")
async def clear_confirm(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text("⚠️ Удалить все сделки реальной торговли?", reply_markup=confirm_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "clear_yes")
async def clear_yes(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    clear_trades(uid)
    await call.message.edit_text(get_text(lang, "cleared"), reply_markup=real_menu_kb(lang))
    await call.answer()

# ---------- БЭКТЕСТ ----------
@dp.callback_query(F.data == "backtest_add_trade")
async def backtest_add_trade(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(BacktestForm.period_start)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "bt_period_start"), reply_markup=back_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "backtest_list_trades")
async def backtest_list_trades(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_backtests(uid)
    if df.empty:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    text = "📋 **Список бэктестов**\n\n"
    for _, row in df.iterrows():
        pnl = row['pnl_usd']
        emoji = "✅" if pnl > 0 else ("❌" if pnl < 0 else "⚖️")
        pnl_text = f"+${pnl:.0f}" if pnl > 0 else (f"${pnl:.0f}" if pnl < 0 else "БУ")
        text += f"{emoji} #{row['id']} {row['asset']} | {row['direction']} | {pnl_text} ({row['pnl_r']:.1f}R) | {row['period_start']}\n"
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=backtest_menu_kb(lang))
    await call.answer()

@dp.message(BacktestForm.period_start)
async def bt_start(msg: Message, state: FSMContext):
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_start=d)
        await state.set_state(BacktestForm.period_end)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "bt_period_end"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_date"))

@dp.message(BacktestForm.period_end)
async def bt_end(msg: Message, state: FSMContext):
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_end=d)
        await state.set_state(BacktestForm.timeframe)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "bt_timeframe"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_date"))

@dp.message(BacktestForm.timeframe)
async def bt_tf(msg: Message, state: FSMContext):
    await state.update_data(timeframe=msg.text.upper())
    await state.set_state(BacktestForm.commission)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "bt_commission"), reply_markup=back_kb(lang))

@dp.message(BacktestForm.commission)
async def bt_comm(msg: Message, state: FSMContext):
    try:
        await state.update_data(commission=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.spread)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "bt_spread"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.spread)
async def bt_spread(msg: Message, state: FSMContext):
    try:
        await state.update_data(spread=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.asset)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.asset)
async def bt_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(BacktestForm.direction)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang, "backtest"))

@dp.callback_query(F.data.startswith("backtest_LONG") | F.data.startswith("backtest_SHORT"))
async def bt_direction(call: CallbackQuery, state: FSMContext):
    direction = call.data.split("_")[1]
    await state.update_data(direction=direction)
    await state.set_state(BacktestForm.entry_price)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(BacktestForm.entry_price)
async def bt_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.sl_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_sl"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.sl_price)
async def bt_sl(msg: Message, state: FSMContext):
    try:
        await state.update_data(sl_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.tp_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_tp"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.tp_price)
async def bt_tp(msg: Message, state: FSMContext):
    try:
        await state.update_data(tp_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.exit_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_exit_price_bt"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.exit_price)
async def bt_exit(msg: Message, state: FSMContext):
    try:
        exit_p = float(msg.text.replace(",", "."))
        data = await state.get_data()
        direction = data['direction']
        entry = data['entry_price']
        sl = data['sl_price']
        comm = data['commission']/100
        spread = data['spread']
        if direction == "LONG":
            pnl = (exit_p - entry) - comm*entry - spread
        else:
            pnl = (entry - exit_p) - comm*entry - spread
        risk = abs(entry - sl)
        r = pnl/risk if risk else 0
        await state.update_data(exit_price=exit_p, pnl_usd=pnl, pnl_r=r)
        await state.set_state(BacktestForm.entry_time)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(f"📊 P&L: ${pnl:.2f} ({r:.2f}R)\n\n{get_text(lang, 'enter_entry_time')}", reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.entry_time)
async def bt_etime(msg: Message, state: FSMContext):
    await state.update_data(entry_time=msg.text)
    await state.set_state(BacktestForm.exit_time)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_exit_time"), reply_markup=back_kb(lang))

@dp.message(BacktestForm.exit_time)
async def bt_xtime(msg: Message, state: FSMContext):
    await state.update_data(exit_time=msg.text)
    await state.set_state(BacktestForm.setup)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_setup"), reply_markup=back_kb(lang))

@dp.message(BacktestForm.setup)
async def bt_setup(msg: Message, state: FSMContext):
    await state.update_data(setup=msg.text)
    await state.set_state(BacktestForm.trigger)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_trigger"), reply_markup=back_kb(lang))

@dp.message(BacktestForm.trigger)
async def bt_trigger(msg: Message, state: FSMContext):
    await state.update_data(trigger=msg.text)
    await state.set_state(BacktestForm.signal_quality)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_quality"), reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1", callback_data="q_1"), InlineKeyboardButton(text="2", callback_data="q_2"),
         InlineKeyboardButton(text="3", callback_data="q_3"), InlineKeyboardButton(text="4", callback_data="q_4"),
         InlineKeyboardButton(text="5", callback_data="q_5")]
    ]))

@dp.callback_query(F.data.startswith("q_"))
async def bt_quality(call: CallbackQuery, state: FSMContext):
    q = int(call.data.split("_")[1])
    await state.update_data(signal_quality=q)
    await state.set_state(BacktestForm.link_chart)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_link_bt"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(BacktestForm.link_chart)
async def bt_link(msg: Message, state: FSMContext):
    link = msg.text if msg.text != "0" else "-"
    data = await state.get_data()
    bt_data = {
        'user_id': msg.from_user.id,
        'period_start': data['period_start'],
        'period_end': data['period_end'],
        'timeframe': data['timeframe'],
        'commission': data['commission'],
        'spread': data['spread'],
        'asset': data['asset'],
        'direction': data['direction'],
        'entry_price': data['entry_price'],
        'exit_price': data['exit_price'],
        'sl_price': data['sl_price'],
        'tp_price': data['tp_price'],
        'pnl_usd': data['pnl_usd'],
        'pnl_r': data['pnl_r'],
        'signal_quality': data['signal_quality'],
        'setup': data['setup'],
        'trigger': data['trigger'],
        'link_chart': link,
        'entry_time': data['entry_time'],
        'exit_time': data['exit_time']
    }
    save_backtest(bt_data)
    await state.clear()
    lang = get_user_lang(msg.from_user.id)
    await msg.answer("✅ Бэктест сохранён!", reply_markup=backtest_menu_kb(lang))

# ---------- КОМАНДЫ ----------
@dp.message(Command("new"))
async def cmd_new(msg: Message, state: FSMContext):
    await state.clear()
    await state.set_state(RealTradeForm.asset)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))

@dp.message(Command("stats"))
async def cmd_stats(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid)
    text = get_real_stats_text(df, lang, "stats_header")
    await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("day"))
async def cmd_day(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, date_filter="day")
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    text = get_real_stats_text(df, lang, "stats_today")
    await msg.answer(f"📆 **Статистика за сегодня**\n\n{text}", parse_mode="Markdown")

@dp.message(Command("week"))
async def cmd_week(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, date_filter="week")
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    text = get_real_stats_text(df, lang, "stats_week")
    await msg.answer(f"📅 **Статистика за неделю**\n\n{text}", parse_mode="Markdown")

@dp.message(Command("month"))
async def cmd_month(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, date_filter="month")
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    text = get_real_stats_text(df, lang, "stats_month")
    await msg.answer(f"📊 **Статистика за месяц**\n\n{text}", parse_mode="Markdown")

@dp.message(Command("clear"))
async def cmd_clear(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    clear_trades(uid)
    await msg.answer(get_text(lang, "cleared"))

@dp.message(Command("get_real"))
async def cmd_get_real(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    fname = export_real_to_excel(df, uid)
    await msg.answer_document(document=FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    os.remove(fname)

@dp.message(Command("get_backtest"))
async def cmd_get_backtest(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_backtests(uid)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    fname = export_backtest_to_excel(df, uid)
    await msg.answer_document(document=FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    os.remove(fname)

# ========== ЗАПУСК ==========
async def set_commands(bot: Bot):
    await bot.set_my_commands([
        BotCommand(command="start", description="Главное меню"),
        BotCommand(command="new", description="➕ Новая сделка"),
        BotCommand(command="stats", description="Вся статистика"),
        BotCommand(command="day", description="Статистика за день"),
        BotCommand(command="week", description="Статистика за неделю"),
        BotCommand(command="month", description="Статистика за месяц"),
        BotCommand(command="clear", description="Очистить журнал"),
        BotCommand(command="get_real", description="Excel (реальная торговля)"),
        BotCommand(command="get_backtest", description="Excel (бэктест)"),
    ])

async def main():
    global bot
    init_dbs()
    bot = Bot(token=BOT_TOKEN)
    await set_commands(bot)
    print("✅ Бот успешно запущен! Все функции восстановлены.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
