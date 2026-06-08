import asyncio
import sqlite3
import os
from datetime import datetime, timedelta
from aiohttp import web

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, CallbackQuery, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton, BotCommand
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ========== КОНФИГ ==========
BOT_TOKEN = "8584035526:AAG8Q15ym8TONEAOH4_8_eQaXnsV4VhhIYs"
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
        "error_date": "❌ Ошибка! Введите дату в формате ДД.ММ.ГГГГ."
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
        "error_date": "❌ Error! Enter date in DD.MM.YYYY format."
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
def get_stats_text(df, lang):
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
    return (
        f"{get_text(lang, 'stats_header')}\n\n"
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
        f"⚙️ Профит-фактор: {pf:.2f}"
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

def trades_list_kb(trades, page, total_pages, lang):
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
        buttons.append([InlineKeyboardButton(text=f"{row['asset']} {emoji} {pnl_text}", callback_data=f"trade_{row['id']}")])
    nav = []
    if page > 1:
        nav.append(InlineKeyboardButton(text=get_text(lang, "prev"), callback_data=f"trades_page_{page-1}"))
    if page < total_pages:
        nav.append(InlineKeyboardButton(text=get_text(lang, "next"), callback_data=f"trades_page_{page+1}"))
    if nav:
        buttons.append(nav)
    buttons.append([InlineKeyboardButton(text=get_text(lang, "refresh"), callback_data="real_list_trades"),
                   InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

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
    asset = State()
    direction = State()
    entry_price = State()
    exit_price = State()
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
async def change_lang_menu(call: CallbackQuery):
    await call.message.edit_text(get_text("ru", "select_language"), reply_markup=lang_kb())
    await call.answer()

@dp.callback_query(F.data == "support")
async def support_menu(call: CallbackQuery):
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
async def link_yes_handler(call: CallbackQuery, state: FSMContext):
    await state.set_state(RealTradeForm.link_url)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_link"), reply_markup=back_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "link_no")
async def link_no_handler(call: CallbackQuery, state: FSMContext):
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
    df = get_trades_filtered(uid)
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
    await state.update_data(page=page)
    text = get_text(lang, "recent_trades", page=page, total_pages=pages)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=trades_list_kb(trades_df, page, pages, lang))
    await call.answer()

@dp.callback_query(F.data.startswith("trades_page_"))
async def trades_page(call: CallbackQuery, state: FSMContext):
    page = int(call.data.split("_")[2])
    await state.update_data(page=page)
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

# ---------- СТАТИСТИКА ----------
@dp.callback_query(F.data == "stats_menu")
async def stats_menu(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    df = get_trades_filtered(call.from_user.id)
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]]))
    await call.answer()

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
    await call.message.edit_text("⚠️ Удалить все сделки?", reply_markup=confirm_kb(lang))
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
    await state.set_state(BacktestForm.asset)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))

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
        await state.update_data(exit_price=exit_p)
        await state.set_state(BacktestForm.link_chart)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_link_bt"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(BacktestForm.link_chart)
async def bt_link(msg: Message, state: FSMContext):
    link = msg.text if msg.text != "0" else "-"
    data = await state.get_data()
    bt_data = {
        'user_id': msg.from_user.id,
        'period_start': data['period_start'],
        'period_end': data['period_end'],
        'timeframe': data['timeframe'],
        'commission': 0,
        'spread': 0,
        'asset': data['asset'],
        'direction': data['direction'],
        'entry_price': data['entry_price'],
        'exit_price': data['exit_price'],
        'sl_price': 0,
        'tp_price': 0,
        'pnl_usd': 0,
        'pnl_r': 0,
        'signal_quality': 3,
        'setup': '-',
        'trigger': '-',
        'link_chart': link,
        'entry_time': '00:00',
        'exit_time': '00:00'
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
    text = get_stats_text(df, lang)
    await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("day"))
async def cmd_day(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, date_filter="day")
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    text = get_stats_text(df, lang)
    await msg.answer(f"📆 **Статистика за сегодня**\n\n{text}", parse_mode="Markdown")

@dp.message(Command("week"))
async def cmd_week(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, date_filter="week")
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    text = get_stats_text(df, lang)
    await msg.answer(f"📅 **Статистика за неделю**\n\n{text}", parse_mode="Markdown")

@dp.message(Command("month"))
async def cmd_month(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades_filtered(uid, date_filter="month")
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    text = get_stats_text(df, lang)
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

# ========== ВЕБ-СЕРВЕР ДЛЯ RENDER ==========
async def health_check(request):
    return web.Response(text="Bot is alive!")

async def run_web_server():
    app = web.Application()
    app.router.add_get('/', health_check)
    app.router.add_get('/health', health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()
    print(f"🌐 Web server started on port {port}")
    # Keep the server running
    await asyncio.Event().wait()

async def set_commands(bot: Bot):
    await bot.set_my_commands([
        BotCommand(command="start", description="Main menu"),
        BotCommand(command="new", description="➕ New trade"),
        BotCommand(command="stats", description="📊 All statistics"),
        BotCommand(command="day", description="📆 Today's statistics"),
        BotCommand(command="week", description="📅 Weekly statistics"),
        BotCommand(command="month", description="📊 Monthly statistics"),
        BotCommand(command="clear", description="🗑 Clear journal"),
        BotCommand(command="get_real", description="📎 Excel (Real trading)"),
        BotCommand(command="get_backtest", description="📎 Excel (Backtest)"),
    ])

async def main():
    global bot
    init_dbs()
    bot = Bot(token=BOT_TOKEN)
    await set_commands(bot)
    print("✅ Bot started successfully!")
    # Run web server and bot together
    await asyncio.gather(
        run_web_server(),
        dp.start_polling(bot)
    )

if __name__ == "__main__":
    asyncio.run(main())
