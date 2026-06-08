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
BOT_TOKEN = "8584035526:AAG8Q15ym8TONEAOH4_8_eQaXnsV4VhhIYs"
DB_NAME = "trades.db"
BT_DB_NAME = "backtests.db"

# ========== ЛОКАЛИЗАЦИЯ ==========
TEXTS = {
    "ru": {
        "select_mode": "🎛 **Выберите режим работы:**",
        "mode_real": "📊 Реальная торговля",
        "mode_backtest": "🔄 Бэктест",
        "add_trade": "➕ Сделка",
        "stats": "📊 Статистика",
        "excel_real": "📎 Excel (Реал)",
        "excel_backtest": "📎 Excel (Бэктест)",
        "clear": "🗑 Очистить",
        "settings": "⚙️ Настройки",
        "support": "📞 Поддержка",
        "back": "🔙 Назад",
        "yes": "✅ Да",
        "no": "❌ Нет",
        "confirm_clear": "⚠️ ДА, УДАЛИТЬ ВСЁ",
        "cleared": "🗑 Журнал очищен.",
        "no_data": "📭 Нет данных.",
        "no_data_add_trade": "📭 Нет данных. Добавьте первую сделку через меню ➕ Сделка.",
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
        "bt_period_start": "📅 Введите НАЧАЛО периода (ДД.ММ.ГГГГ):",
        "bt_period_end": "📅 Введите КОНЕЦ периода (ДД.ММ.ГГГГ):",
        "bt_timeframe": "⏱ Введите таймфрейм (M5, H1, H4, D1, W1):",
        "bt_commission": "💸 Комиссия в % (0.1):",
        "bt_spread": "📊 Спред в пунктах (1.5):",
        "enter_sl": "🛑 Введите цену Stop-Loss:",
        "enter_tp": "🎯 Введите цену Take-Profit (0 если нет):",
        "enter_exit_price_bt": "💰 Введите цену выхода:",
        "enter_entry_time": "⏰ Время входа (ЧЧ:ММ):",
        "enter_exit_time": "⏰ Время выхода (ЧЧ:ММ):",
        "enter_setup": "🎯 Название сетапа (стратегии):",
        "enter_trigger": "⚡ Триггер (сигнал):",
        "enter_quality": "⭐ Оцените качество сигнала (1-5):",
        "enter_link_bt": "🔗 Ссылка на скриншот (0 если нет):",
        "error_number": "❌ Ошибка! Введите число.",
        "error_date": "❌ Ошибка! Введите дату в формате ДД.ММ.ГГГГ.",
        "select_language": "🌐 Выберите язык:",
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
        "stats_header": "📊 Ваша статистика",
        "stats_today": "📆 Статистика за сегодня",
        "stats_week": "📅 Статистика за неделю",
        "stats_month": "📊 Статистика за месяц",
        "total_trades": "📋 Всего сделок: {total}",
        "wins": "✅ Прибыльных: {wins}",
        "losses": "❌ Убыточных: {losses}",
        "winrate": "🎯 Винрейт: {wr:.1f}%",
        "longs_shorts": "📈 Лонги: {longs} | 📉 Шорты: {shorts}",
        "total_pnl": "💰 Суммарный P&L: ${total_pnl:.2f}",
        "avg_pnl": "📊 Средняя сделка: ${avg_pnl:.2f}",
        "best": "🏆 Лучшая: +${best:.2f}",
        "worst": "💀 Худшая: ${worst:.2f}",
        "pf": "⚙️ Профит-фактор: {pf:.2f}",
        "excel_ready": "📊 Ваш отчёт",
        "trade_detail": "📋 **Сделка #{id}**\n\nАктив: {asset}\nНаправление: {direction}\nВход: ${entry}\nВыход: ${exit}\nОбъём: {volume}\nP&L: ${pnl}\nИсход: {result}\nДата: {date}\nСсылки:\n{links}",
        "recent_trades": "🕒 **Недавние сделки** (последние {count}):",
        "sort_newest": "📅 Сначала новые",
        "sort_oldest": "📅 Сначала старые"
    },
    "en": {
        "select_mode": "🎛 **Select mode:**",
        "mode_real": "📊 Real Trading",
        "mode_backtest": "🔄 Backtest",
        "add_trade": "➕ Add Trade",
        "stats": "📊 Statistics",
        "excel_real": "📎 Excel (Real)",
        "excel_backtest": "📎 Excel (Backtest)",
        "clear": "🗑 Clear",
        "settings": "⚙️ Settings",
        "support": "📞 Support",
        "back": "🔙 Back",
        "yes": "✅ Yes",
        "no": "❌ No",
        "confirm_clear": "⚠️ YES, DELETE ALL",
        "cleared": "🗑 Journal cleared.",
        "no_data": "📭 No data.",
        "no_data_add_trade": "📭 No data. Add your first trade via ➕ Add Trade.",
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
        "bt_period_start": "📅 Enter START date (DD.MM.YYYY):",
        "bt_period_end": "📅 Enter END date (DD.MM.YYYY):",
        "bt_timeframe": "⏱ Enter timeframe (M5, H1, H4, D1, W1):",
        "bt_commission": "💸 Commission % (0.1):",
        "bt_spread": "📊 Spread in points (1.5):",
        "enter_sl": "🛑 Enter Stop-Loss price:",
        "enter_tp": "🎯 Enter Take-Profit price (0 if none):",
        "enter_exit_price_bt": "💰 Enter exit price:",
        "enter_entry_time": "⏰ Entry time (HH:MM):",
        "enter_exit_time": "⏰ Exit time (HH:MM):",
        "enter_setup": "🎯 Setup name (strategy):",
        "enter_trigger": "⚡ Trigger (signal):",
        "enter_quality": "⭐ Rate signal quality (1-5):",
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
        "stats_header": "📊 Your statistics",
        "stats_today": "📆 Today's statistics",
        "stats_week": "📅 Weekly statistics",
        "stats_month": "📊 Monthly statistics",
        "total_trades": "📋 Total trades: {total}",
        "wins": "✅ Winning: {wins}",
        "losses": "❌ Losing: {losses}",
        "winrate": "🎯 Win rate: {wr:.1f}%",
        "longs_shorts": "📈 Longs: {longs} | 📉 Shorts: {shorts}",
        "total_pnl": "💰 Total P&L: ${total_pnl:.2f}",
        "avg_pnl": "📊 Average trade: ${avg_pnl:.2f}",
        "best": "🏆 Best: +${best:.2f}",
        "worst": "💀 Worst: ${worst:.2f}",
        "pf": "⚙️ Profit factor: {pf:.2f}",
        "excel_ready": "📊 Your report",
        "trade_detail": "📋 **Trade #{id}**\n\nAsset: {asset}\nDirection: {direction}\nEntry: ${entry}\nExit: ${exit}\nVolume: {volume}\nP&L: ${pnl}\nOutcome: {result}\nDate: {date}\nLinks:\n{links}",
        "recent_trades": "🕒 **Recent trades** (last {count}):",
        "sort_newest": "📅 Newest first",
        "sort_oldest": "📅 Oldest first"
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

def get_trades(user_id, start_date=None, asset=None, sort_by_date="DESC"):
    conn = sqlite3.connect(DB_NAME)
    query = "SELECT * FROM trades WHERE user_id = ?"
    params = [user_id]
    if start_date:
        query += " AND trade_date >= ?"
        params.append(start_date)
    if asset:
        query += " AND asset = ?"
        params.append(asset)
    query += f" ORDER BY trade_date {sort_by_date}"
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
    df_exp = df_exp[[
        'trade_date', 'asset', 'direction', 'entry_price', 'exit_price', 'volume', 'pnl', 'result', 'comment', 'links', 'emotion'
    ]]
    df_exp.columns = ['📅 Дата', '🪙 Актив', '📈 Направление', '💰 Вход', '💰 Выход', '📊 Объём', '💵 P&L', '🎯 Исход', '📝 Комментарий', '🔗 Ссылки', '😊 Эмоции']
    df_exp['📈 Направление'] = df_exp['📈 Направление'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_exp['🎯 Исход'] = df_exp['🎯 Исход'].replace({'TAKE': '✅ Тейк', 'STOP': '❌ Стоп'})
    df_exp = df_exp.sort_values('📅 Дата', ascending=False)
    fname = f"real_{user_id}.xlsx"
    with pd.ExcelWriter(fname, engine='openpyxl') as w:
        df_exp.to_excel(w, sheet_name='Реальная торговля', index=False)
        ws = w.sheets['Реальная торговля']
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2b6cb0", end_color="2b6cb0", fill_type="solid")
        for col in range(1, len(df_exp.columns)+1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
        green = PatternFill(start_color="c6f7d0", end_color="c6f7d0", fill_type="solid")
        red = PatternFill(start_color="fecaca", end_color="fecaca", fill_type="solid")
        for row in range(2, len(df_exp)+2):
            val = ws.cell(row=row, column=7).value
            if val and val > 0:
                for c in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=c).fill = green
            elif val and val < 0:
                for c in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=c).fill = red
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
    df_exp = df_exp[[
        'period_start', 'period_end', 'timeframe', 'asset', 'direction',
        'entry_price', 'exit_price', 'pnl_usd', 'pnl_r', 'signal_quality', 'setup', 'trigger'
    ]]
    df_exp.columns = ['📅 Начало', '📅 Конец', '⏱ Таймфрейм', '🪙 Актив', '📈 Направление',
                      '💰 Вход', '💰 Выход', '💵 P&L', '📊 P&L (R)', '⭐ Качество', '🎯 Сетап', '⚡ Триггер']
    df_exp['📈 Направление'] = df_exp['📈 Направление'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_exp = df_exp.sort_values('📅 Начало', ascending=False)
    fname = f"backtest_{user_id}.xlsx"
    with pd.ExcelWriter(fname, engine='openpyxl') as w:
        df_exp.to_excel(w, sheet_name='Бэктест', index=False)
        ws = w.sheets['Бэктест']
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2b6cb0", end_color="2b6cb0", fill_type="solid")
        for col in range(1, len(df_exp.columns)+1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
        green = PatternFill(start_color="c6f7d0", end_color="c6f7d0", fill_type="solid")
        red = PatternFill(start_color="fecaca", end_color="fecaca", fill_type="solid")
        for row in range(2, len(df_exp)+2):
            val = ws.cell(row=row, column=8).value
            if val and val > 0:
                for c in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=c).fill = green
            elif val and val < 0:
                for c in range(1, len(df_exp.columns)+1):
                    ws.cell(row=row, column=c).fill = red
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
        f"{get_text(lang, 'total_trades', total=total)}\n"
        f"{get_text(lang, 'wins', wins=wins)}\n"
        f"{get_text(lang, 'losses', losses=losses)}\n"
        f"{get_text(lang, 'winrate', wr=wr)}\n"
        f"{get_text(lang, 'longs_shorts', longs=longs, shorts=shorts)}\n"
        f"{get_text(lang, 'total_pnl', total_pnl=total_pnl)}\n"
        f"{get_text(lang, 'avg_pnl', avg_pnl=avg_pnl)}\n"
        f"{get_text(lang, 'best', best=best)}\n"
        f"{get_text(lang, 'worst', worst=worst)}\n"
        f"{get_text(lang, 'pf', pf=pf)}"
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
        [InlineKeyboardButton(text=get_text(lang, "add_trade"), callback_data="add_real_trade")],
        [InlineKeyboardButton(text=get_text(lang, "stats"), callback_data="stats_menu")],
        [InlineKeyboardButton(text=get_text(lang, "excel_real"), callback_data="get_real_excel")],
        [InlineKeyboardButton(text=get_text(lang, "clear"), callback_data="clear_confirm")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu"),
         InlineKeyboardButton(text=get_text(lang, "support"), callback_data="support")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def backtest_menu_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "add_trade"), callback_data="add_backtest_trade")],
        [InlineKeyboardButton(text=get_text(lang, "stats"), callback_data="stats_menu")],
        [InlineKeyboardButton(text=get_text(lang, "excel_backtest"), callback_data="get_backtest_excel")],
        [InlineKeyboardButton(text=get_text(lang, "clear"), callback_data="clear_confirm")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu"),
         InlineKeyboardButton(text=get_text(lang, "support"), callback_data="support")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def direction_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "long"), callback_data="LONG"),
         InlineKeyboardButton(text=get_text(lang, "short"), callback_data="SHORT")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def result_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "take"), callback_data="TAKE"),
         InlineKeyboardButton(text=get_text(lang, "stop"), callback_data="STOP")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def emotion_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "emotion_calm"), callback_data="emotion_calm")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_fear"), callback_data="emotion_fear")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_greed"), callback_data="emotion_greed")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_tilt"), callback_data="emotion_tilt")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_confidence"), callback_data="emotion_confidence")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def yesno_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "yes"), callback_data="yes")],
        [InlineKeyboardButton(text=get_text(lang, "no"), callback_data="no")]
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

def back_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
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

# ========== РОУТЕР ==========
bot = None
dp = Dispatcher()

# ========== КОМАНДЫ ==========
@dp.message(CommandStart())
async def start_cmd(msg: Message, state: FSMContext):
    await state.clear()
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    if not lang:
        await msg.answer(get_text("ru", "select_language"), reply_markup=lang_kb())
        return
    await msg.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))

@dp.message(Command("stats"))
async def cmd_stats(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid)
    text = get_stats_text(df, lang)
    await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("day"))
async def cmd_day(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    today = datetime.now().strftime("%Y-%m-%d")
    df = get_trades(uid, start_date=today)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    total = len(df)
    wins = len(df[df['pnl'] > 0])
    losses = len(df[df['pnl'] < 0])
    wr = wins/total*100 if total else 0
    longs = len(df[df['direction'] == 'LONG'])
    shorts = len(df[df['direction'] == 'SHORT'])
    total_pnl = df['pnl'].sum()
    await msg.answer(
        f"📆 **Статистика за сегодня**\n\n"
        f"📋 Сделок: {total}\n"
        f"✅ Прибыльных: {wins}\n"
        f"❌ Убыточных: {losses}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"📈 Лонги: {longs} | 📉 Шорты: {shorts}\n"
        f"💰 P&L: ${total_pnl:.2f}",
        parse_mode="Markdown"
    )

@dp.message(Command("week"))
async def cmd_week(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    week_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    df = get_trades(uid, start_date=week_ago)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    total = len(df)
    wins = len(df[df['pnl'] > 0])
    losses = len(df[df['pnl'] < 0])
    wr = wins/total*100 if total else 0
    longs = len(df[df['direction'] == 'LONG'])
    shorts = len(df[df['direction'] == 'SHORT'])
    total_pnl = df['pnl'].sum()
    await msg.answer(
        f"📅 **Статистика за неделю**\n\n"
        f"📋 Сделок: {total}\n"
        f"✅ Прибыльных: {wins}\n"
        f"❌ Убыточных: {losses}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"📈 Лонги: {longs} | 📉 Шорты: {shorts}\n"
        f"💰 P&L: ${total_pnl:.2f}",
        parse_mode="Markdown"
    )

@dp.message(Command("month"))
async def cmd_month(msg: Message):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    month_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    df = get_trades(uid, start_date=month_ago)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    total = len(df)
    wins = len(df[df['pnl'] > 0])
    losses = len(df[df['pnl'] < 0])
    wr = wins/total*100 if total else 0
    longs = len(df[df['direction'] == 'LONG'])
    shorts = len(df[df['direction'] == 'SHORT'])
    total_pnl = df['pnl'].sum()
    await msg.answer(
        f"📊 **Статистика за месяц**\n\n"
        f"📋 Сделок: {total}\n"
        f"✅ Прибыльных: {wins}\n"
        f"❌ Убыточных: {losses}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"📈 Лонги: {longs} | 📉 Шорты: {shorts}\n"
        f"💰 P&L: ${total_pnl:.2f}",
        parse_mode="Markdown"
    )

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
    df = get_trades(uid)
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

# ========== НАСТРОЙКИ ==========
@dp.callback_query(F.data.startswith("lang_"))
async def set_lang(call: CallbackQuery, state: FSMContext):
    lang = call.data.split("_")[1]
    set_user_lang(call.from_user.id, lang)
    await call.message.delete()
    await call.message.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "settings_menu")
async def settings(call: CallbackQuery):
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

@dp.callback_query(F.data == "back_mode")
async def back_mode(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

# ========== ВЫБОР РЕЖИМА ==========
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

# ========== ДОБАВЛЕНИЕ РЕАЛЬНОЙ СДЕЛКИ ==========
@dp.callback_query(F.data == "add_real_trade")
async def add_real_trade(call: CallbackQuery, state: FSMContext):
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
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang))

@dp.callback_query(F.data.in_(["LONG", "SHORT"]))
async def real_direction(call: CallbackQuery, state: FSMContext):
    await state.update_data(direction=call.data)
    await state.set_state(RealTradeForm.entry_price)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(RealTradeForm.entry_price)
async def real_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.exit_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_exit_price"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(RealTradeForm.exit_price)
async def real_exit(msg: Message, state: FSMContext):
    try:
        await state.update_data(exit_price=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.volume)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_volume"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.message(RealTradeForm.volume)
async def real_volume(msg: Message, state: FSMContext):
    try:
        await state.update_data(volume=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.result)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "choose_result"), reply_markup=result_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@dp.callback_query(F.data.in_(["TAKE", "STOP"]))
async def real_result(call: CallbackQuery, state: FSMContext):
    await state.update_data(result=call.data)
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
    await msg.answer(get_text(lang, "add_link_question"), reply_markup=yesno_kb(lang))

@dp.callback_query(F.data == "yes")
async def real_add_link(call: CallbackQuery, state: FSMContext):
    await state.set_state(RealTradeForm.link_url)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_link"), reply_markup=back_kb(lang))
    await call.answer()

@dp.callback_query(F.data == "no")
async def real_skip_links(call: CallbackQuery, state: FSMContext):
    await state.update_data(links="")
    await state.set_state(RealTradeForm.trade_date)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_date"), reply_markup=back_kb(lang))
    await call.answer()

@dp.message(RealTradeForm.link_url)
async def real_get_link(msg: Message, state: FSMContext):
    url = msg.text
    await state.update_data(link_url=url)
    await state.set_state(RealTradeForm.link_tf)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_timeframe"), reply_markup=back_kb(lang))

@dp.message(RealTradeForm.link_tf)
async def real_get_tf(msg: Message, state: FSMContext):
    tf = msg.text
    data = await state.get_data()
    links = data.get("links", "")
    new_link = f"{tf}: {data.get('link_url')}"
    if links:
        links += f"\n{new_link}"
    else:
        links = new_link
    await state.update_data(links=links)
    await state.set_state(RealTradeForm.add_link)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "link_saved"), reply_markup=yesno_kb(lang))

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
    await msg.answer(get_text(lang, "enter_emotion"), reply_markup=emotion_kb(lang))

@dp.callback_query(F.data.startswith("emotion_"))
async def real_emotion(call: CallbackQuery, state: FSMContext):
    emotion_map = {
        "emotion_calm": "😊 Спокойствие",
        "emotion_fear": "😨 Страх",
        "emotion_greed": "😈 Жадность",
        "emotion_tilt": "🤬 Тильт",
        "emotion_confidence": "😌 Уверенность"
    }
    emotion = emotion_map.get(call.data, "😊 Спокойствие")
    data = await state.get_data()
    
    direction = data['direction']
    entry = data['entry_price']
    exit_p = data['exit_price']
    vol = data['volume']
    if direction == "LONG":
        pnl = (exit_p - entry) * vol
    else:
        pnl = (entry - exit_p) * vol
    
    save_trade(
        user_id=call.from_user.id,
        asset=data['asset'],
        direction=data['direction'],
        entry_price=data['entry_price'],
        exit_price=data['exit_price'],
        volume=data['volume'],
        pnl=pnl,
        result=data['result'],
        comment=data['comment'],
        trade_date=data['trade_date'],
        links=data.get('links', ''),
        emotion=emotion
    )
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "trade_saved"), parse_mode="Markdown", reply_markup=real_menu_kb(lang))
    await call.answer()

# ========== СТАТИСТИКА МЕНЮ ==========
@dp.callback_query(F.data == "stats_menu")
async def stats_menu(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "stats_menu"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "stats_all"), callback_data="stats_all")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_asset"), callback_data="stats_by_asset")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_date"), callback_data="stats_by_date")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_emotion"), callback_data="stats_by_emotion")],
        [InlineKeyboardButton(text=get_text(lang, "stats_recent"), callback_data="stats_recent")],
        [InlineKeyboardButton(text=get_text(lang, "stats_sort"), callback_data="stats_sort")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ]))
    await call.answer()

@dp.callback_query(F.data == "stats_all")
async def stats_all(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid)
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]]))
    await call.answer()

@dp.callback_query(F.data == "stats_by_asset")
async def stats_by_asset(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    assets = get_all_assets(uid)
    if not assets:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    buttons = [[InlineKeyboardButton(text=a, callback_data=f"asset_{a}")] for a in assets]
    buttons.append([InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")])
    await call.message.edit_text(get_text(lang, "select_asset"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@dp.callback_query(F.data.startswith("asset_"))
async def show_asset_stats(call: CallbackQuery):
    asset = call.data.split("_")[1]
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid, asset=asset)
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_by_asset")]]))
    await call.answer()

@dp.callback_query(F.data == "stats_by_date")
async def stats_by_date(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_period"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "stats_today"), callback_data="period_day")],
        [InlineKeyboardButton(text=get_text(lang, "stats_week"), callback_data="period_week")],
        [InlineKeyboardButton(text=get_text(lang, "stats_month"), callback_data="period_month")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ]))
    await call.answer()

@dp.callback_query(F.data == "period_day")
async def period_day(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    today = datetime.now().strftime("%Y-%m-%d")
    df = get_trades(uid, start_date=today)
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_by_date")]]))
    await call.answer()

@dp.callback_query(F.data == "period_week")
async def period_week(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    week_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    df = get_trades(uid, start_date=week_ago)
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_by_date")]]))
    await call.answer()

@dp.callback_query(F.data == "period_month")
async def period_month(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    month_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    df = get_trades(uid, start_date=month_ago)
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_by_date")]]))
    await call.answer()

@dp.callback_query(F.data == "stats_by_emotion")
async def stats_by_emotion(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    emotions = ['😊 Спокойствие', '😨 Страх', '😈 Жадность', '🤬 Тильт', '😌 Уверенность']
    df = get_trades(uid)
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
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]]))
    await call.answer()

@dp.callback_query(F.data == "stats_recent")
async def stats_recent(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid, sort_by_date="DESC")
    if df.empty:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    df = df.head(10)
    text = get_text(lang, "recent_trades", count=len(df)) + "\n\n"
    for _, row in df.iterrows():
        emoji = "✅" if row['pnl'] > 0 else "❌"
        text += f"{emoji} #{row['id']} {row['asset']} | {row['direction']} | ${row['pnl']:.0f} | {row['trade_date']}\n"
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]]))
    await call.answer()

@dp.callback_query(F.data == "stats_sort")
async def stats_sort(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "stats_sort"), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "sort_newest"), callback_data="sort_newest")],
        [InlineKeyboardButton(text=get_text(lang, "sort_oldest"), callback_data="sort_oldest")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ]))
    await call.answer()

@dp.callback_query(F.data == "sort_newest")
async def sort_newest(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid, sort_by_date="DESC")
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]]))
    await call.answer()

@dp.callback_query(F.data == "sort_oldest")
async def sort_oldest(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid, sort_by_date="ASC")
    text = get_stats_text(df, lang)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]]))
    await call.answer()

# ========== EXCEL КНОПКИ ==========
@dp.callback_query(F.data == "get_real_excel")
async def get_real_excel(call: CallbackQuery):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    df = get_trades(uid)
    if df.empty:
        await call.answer(get_text(lang, "no_data_add_trade"), show_alert=True)
        return
    fname = export_real_to_excel(df, uid)
    await call.message.answer_document(document=FSInputFile(fname), caption=get_text(lang, "excel_ready"))
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
    await call.message.answer_document(document=FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    os.remove(fname)
    await call.answer()

# ========== ОЧИСТКА ==========
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

# ========== БЭКТЕСТ (упрощённо) ==========
@dp.callback_query(F.data == "add_backtest_trade")
async def add_backtest(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(BacktestForm.period_start)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "bt_period_start"), reply_markup=back_kb(lang))
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
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang))

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

# ========== ЗАПУСК ==========
async def set_commands(bot: Bot):
    await bot.set_my_commands([
        BotCommand(command="start", description="Главное меню"),
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
    print("✅ Торговый бот запущен! Все команды работают.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
