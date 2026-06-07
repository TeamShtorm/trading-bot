import asyncio
import sqlite3
import os
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher, F, Router
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
        "stats_all": "📈 Вся статистика",
        "stats_day": "📆 День",
        "stats_week": "📅 Неделя",
        "stats_month": "📊 Месяц",
        "excel": "📎 Excel",
        "clear": "🗑 Очистить",
        "settings": "⚙️ Настройки",
        "back": "🔙 Назад",
        "yes": "✅ Да",
        "no": "❌ Нет",
        "confirm_clear": "⚠️ ДА, УДАЛИТЬ ВСЁ",
        "cleared": "🗑 Журнал очищен.",
        "no_data": "📭 Нет данных.",
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
        "select_language": "🌐 Выберите язык / Choose language:",
        "language_set": "✅ Язык установлен: русский",
        "language_set_en": "✅ Language set: English",
        "support_info": "📞 **Поддержка**\n\nПо вопросам пишите: @ваш_username",
        "change_lang": "🌐 Сменить язык",
        "support": "📞 Поддержка",
        "settings_menu": "⚙️ **Настройки**",
        "stats_header": "📊 Ваша статистика",
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
        "bt_stats_header": "📊 Статистика бэктеста",
        "bt_total_trades": "📋 Всего сделок: {total}",
        "bt_wins": "✅ Прибыльных: {wins}",
        "bt_losses": "❌ Убыточных: {losses}",
        "bt_winrate": "🎯 Винрейт: {wr:.1f}%",
        "bt_avg_r": "📊 Средний R: {avg_r:.2f}",
        "bt_total_r": "💰 Суммарный R: {total_r:.2f}",
        "bt_quality": "⭐ Качество сигнала: {q:.1f}/5"
    },
    "en": {
        "select_mode": "🎛 **Select mode:**",
        "mode_real": "📊 Real Trading",
        "mode_backtest": "🔄 Backtest",
        "add_trade": "➕ Trade",
        "stats_all": "📈 All Stats",
        "stats_day": "📆 Day",
        "stats_week": "📅 Week",
        "stats_month": "📊 Month",
        "excel": "📎 Excel",
        "clear": "🗑 Clear",
        "settings": "⚙️ Settings",
        "back": "🔙 Back",
        "yes": "✅ Yes",
        "no": "❌ No",
        "confirm_clear": "⚠️ YES, DELETE ALL",
        "cleared": "🗑 Journal cleared.",
        "no_data": "📭 No data.",
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
        "select_language": "🌐 Choose language:",
        "language_set": "✅ Language set: Russian",
        "language_set_en": "✅ Language set: English",
        "support_info": "📞 **Support**\n\nContact: @your_username",
        "change_lang": "🌐 Change language",
        "support": "📞 Support",
        "settings_menu": "⚙️ **Settings**",
        "stats_header": "📊 Your statistics",
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
        "bt_stats_header": "📊 Backtest statistics",
        "bt_total_trades": "📋 Total trades: {total}",
        "bt_wins": "✅ Winning: {wins}",
        "bt_losses": "❌ Losing: {losses}",
        "bt_winrate": "🎯 Win rate: {wr:.1f}%",
        "bt_avg_r": "📊 Avg R: {avg_r:.2f}",
        "bt_total_r": "💰 Total R: {total_r:.2f}",
        "bt_quality": "⭐ Signal quality: {q:.1f}/5"
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
            user_id INTEGER, mode TEXT, asset TEXT, direction TEXT,
            entry_price REAL, exit_price REAL, volume REAL, pnl REAL,
            result TEXT, comment TEXT, trade_date TEXT, links TEXT
        )
    """)
    conn.close()
    conn = sqlite3.connect(BT_DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS backtests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER, period_start TEXT, period_end TEXT, timeframe TEXT,
            commission REAL, spread REAL, asset TEXT, direction TEXT,
            entry_price REAL, exit_price REAL, sl_price REAL, tp_price REAL,
            pnl_usd REAL, pnl_r REAL, signal_quality INTEGER,
            setup TEXT, trigger TEXT, link_chart TEXT, entry_time TEXT, exit_time TEXT
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
    return r[0] if r else None

def set_user_lang(user_id, lang):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("INSERT OR REPLACE INTO users (user_id, lang) VALUES (?, ?)", (user_id, lang))
    conn.commit()
    conn.close()

def save_trade(user_id, mode, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        INSERT INTO trades (user_id, mode, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, mode, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links))
    conn.commit()
    conn.close()

def get_trades(user_id, mode=None, start_date=None):
    conn = sqlite3.connect(DB_NAME)
    q = "SELECT * FROM trades WHERE user_id = ?"
    p = [user_id]
    if mode:
        q += " AND mode = ?"
        p.append(mode)
    if start_date:
        q += " AND trade_date >= ?"
        p.append(start_date)
    df = pd.read_sql_query(q, conn, params=p)
    conn.close()
    return df

def clear_trades(user_id, mode=None):
    conn = sqlite3.connect(DB_NAME)
    q = "DELETE FROM trades WHERE user_id = ?"
    p = [user_id]
    if mode:
        q += " AND mode = ?"
        p.append(mode)
    conn.execute(q, p)
    conn.commit()
    conn.close()

def save_backtest(data):
    conn = sqlite3.connect(BT_DB_NAME)
    conn.execute("""
        INSERT INTO backtests (
            user_id, period_start, period_end, timeframe, commission, spread,
            asset, direction, entry_price, exit_price, sl_price, tp_price,
            pnl_usd, pnl_r, signal_quality, setup, trigger, link_chart, entry_time, exit_time
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data['user_id'], data['period_start'], data['period_end'], data['timeframe'],
        data['commission'], data['spread'], data['asset'], data['direction'],
        data['entry_price'], data['exit_price'], data['sl_price'], data['tp_price'],
        data['pnl_usd'], data['pnl_r'], data['signal_quality'],
        data['setup'], data['trigger'], data['link_chart'], data['entry_time'], data['exit_time']
    ))
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
        'trade_date', 'asset', 'direction', 'entry_price', 'exit_price', 'volume', 'pnl', 'result', 'comment', 'links'
    ]]
    df_exp.columns = ['📅 Дата', '🪙 Актив', '📈 Направление', '💰 Вход', '💰 Выход', '📊 Объём', '💵 P&L', '🎯 Исход', '📝 Комментарий', '🔗 Ссылки']
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
def get_real_stats_text(df, lang, title_key="stats_header"):
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
        f"{get_text(lang, title_key)}\n\n"
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
        f"{get_text(lang, 'bt_wins', wins=wins)}\n"
        f"{get_text(lang, 'bt_losses', losses=losses)}\n"
        f"{get_text(lang, 'bt_winrate', wr=wr)}\n"
        f"{get_text(lang, 'bt_avg_r', avg_r=avg_r)}\n"
        f"{get_text(lang, 'bt_total_r', total_r=total_r)}\n"
        f"{get_text(lang, 'bt_quality', q=avg_q)}"
    )

# ========== КЛАВИАТУРЫ ==========
def mode_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "mode_real"), callback_data="mode_real")],
        [InlineKeyboardButton(text=get_text(lang, "mode_backtest"), callback_data="mode_backtest")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu")]
    ])

def main_menu_kb(lang, mode):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "add_trade"), callback_data="add_trade")],
        [InlineKeyboardButton(text=get_text(lang, "stats_all"), callback_data="stats_all"),
         InlineKeyboardButton(text=get_text(lang, "stats_day"), callback_data="stats_day"),
         InlineKeyboardButton(text=get_text(lang, "stats_week"), callback_data="stats_week"),
         InlineKeyboardButton(text=get_text(lang, "stats_month"), callback_data="stats_month")],
        [InlineKeyboardButton(text=get_text(lang, "excel"), callback_data="get_excel"),
         InlineKeyboardButton(text=get_text(lang, "clear"), callback_data="clear_confirm"),
         InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu")],
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

def yesno_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "yes"), callback_data="yes"),
         InlineKeyboardButton(text=get_text(lang, "no"), callback_data="no")]
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

def quality_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1", callback_data="q_1"), InlineKeyboardButton(text="2", callback_data="q_2"),
         InlineKeyboardButton(text="3", callback_data="q_3"), InlineKeyboardButton(text="4", callback_data="q_4"),
         InlineKeyboardButton(text="5", callback_data="q_5")]
    ])

def back_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

# ========== FSM ==========
class TradeForm(StatesGroup):
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
router = Router()
bot = None

# ========== СТАРТ ==========
@router.message(CommandStart())
async def start_cmd(msg: Message, state: FSMContext):
    await state.clear()
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    if not lang:
        await msg.answer(get_text("ru", "select_language"), reply_markup=lang_kb())
        return
    await msg.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))

@router.callback_query(F.data.startswith("lang_"))
async def set_lang(call: CallbackQuery, state: FSMContext):
    lang = call.data.split("_")[1]
    set_user_lang(call.from_user.id, lang)
    await call.message.delete()
    await start_cmd(call.message, state)
    await call.answer()

@router.callback_query(F.data == "back_mode")
async def back_mode(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

@router.callback_query(F.data == "settings_menu")
async def settings(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "settings_menu"), parse_mode="Markdown", reply_markup=settings_kb(lang))
    await call.answer()

@router.callback_query(F.data == "change_lang")
async def change_lang(call: CallbackQuery):
    await call.message.edit_text(get_text("ru", "select_language"), reply_markup=lang_kb())
    await call.answer()

@router.callback_query(F.data == "support")
async def support(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "support_info"), parse_mode="Markdown", reply_markup=settings_kb(lang))
    await call.answer()

# ========== ВЫБОР РЕЖИМА ==========
@router.callback_query(F.data.in_(["mode_real", "mode_backtest"]))
async def choose_mode(call: CallbackQuery, state: FSMContext):
    mode = call.data.split("_")[1]
    await state.update_data(mode=mode)
    lang = get_user_lang(call.from_user.id) or "ru"
    if mode == "real":
        await state.set_state(TradeForm.asset)
        await call.message.edit_text(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))
    else:
        await state.set_state(BacktestForm.period_start)
        await call.message.edit_text(get_text(lang, "bt_period_start"), reply_markup=back_kb(lang))
    await call.answer()

# ========== РЕАЛЬНАЯ ТОРГОВЛЯ ==========
@router.message(TradeForm.asset)
async def real_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(TradeForm.direction)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang))

@router.callback_query(F.data.in_(["LONG","SHORT"]))
async def real_dir(call: CallbackQuery, state: FSMContext):
    await state.update_data(direction=call.data)
    await state.set_state(TradeForm.entry_price)
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.entry_price)
async def real_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",",".")))
        await state.set_state(TradeForm.exit_price)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "enter_exit_price"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(TradeForm.exit_price)
async def real_exit(msg: Message, state: FSMContext):
    try:
        await state.update_data(exit_price=float(msg.text.replace(",",".")))
        await state.set_state(TradeForm.volume)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "enter_volume"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(TradeForm.volume)
async def real_vol(msg: Message, state: FSMContext):
    try:
        await state.update_data(volume=float(msg.text.replace(",",".")))
        await state.set_state(TradeForm.result)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "choose_result"), reply_markup=result_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.callback_query(F.data.in_(["TAKE","STOP"]))
async def real_res(call: CallbackQuery, state: FSMContext):
    await state.update_data(result=call.data)
    await state.set_state(TradeForm.comment)
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "enter_comment"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.comment)
async def real_comment(msg: Message, state: FSMContext):
    com = msg.text.strip()
    if com == "-":
        com = ""
    await state.update_data(comment=com)
    await state.set_state(TradeForm.add_link)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "add_link_question"), reply_markup=yesno_kb(lang))

@router.callback_query(F.data == "yes")
async def add_link_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(TradeForm.link_url)
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "enter_link"), reply_markup=back_kb(lang))
    await call.answer()

@router.callback_query(F.data == "no")
async def skip_links(call: CallbackQuery, state: FSMContext):
    await state.update_data(links="")
    await state.set_state(TradeForm.trade_date)
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "enter_date"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.link_url)
async def get_link(msg: Message, state: FSMContext):
    url = msg.text
    await state.update_data(link_url=url)
    await state.set_state(TradeForm.link_tf)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "enter_timeframe"), reply_markup=back_kb(lang))

@router.message(TradeForm.link_tf)
async def get_tf(msg: Message, state: FSMContext):
    tf = msg.text
    data = await state.get_data()
    links = data.get("links", "")
    new_link = f"{tf}: {data.get('link_url')}"
    if links:
        links += f"\n{new_link}"
    else:
        links = new_link
    await state.update_data(links=links)
    await state.set_state(TradeForm.add_link)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "link_saved"), reply_markup=yesno_kb(lang))

@router.message(TradeForm.trade_date)
async def real_date(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id) or "ru"
    dstr = msg.text.strip().lower()
    if dstr in ["сегодня", "today"]:
        tdate = datetime.now().strftime("%Y-%m-%d")
    else:
        try:
            tdate = datetime.strptime(dstr, "%d.%m.%Y").strftime("%Y-%m-%d")
        except:
            await msg.answer(get_text(lang, "error_date"))
            return
    data = await state.get_data()
    direction = data['direction']
    entry = data['entry_price']
    exit_p = data['exit_price']
    vol = data['volume']
    if direction == "LONG":
        pnl = (exit_p - entry) * vol
    else:
        pnl = (entry - exit_p) * vol
    save_trade(msg.from_user.id, "real", data['asset'], direction, entry, exit_p, vol, pnl,
               data['result'], data['comment'], tdate, data.get('links', ''))
    await state.clear()
    mode = data.get('mode', 'real')
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "trade_saved"), reply_markup=main_menu_kb(lang, mode))

# ========== БЭКТЕСТ ==========
@router.message(BacktestForm.period_start)
async def bt_start(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id) or "ru"
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_start=d)
        await state.set_state(BacktestForm.period_end)
        await msg.answer(get_text(lang, "bt_period_end"), reply_markup=back_kb(lang))
    except:
        await msg.answer(get_text(lang, "error_date"))

@router.message(BacktestForm.period_end)
async def bt_end(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id) or "ru"
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_end=d)
        await state.set_state(BacktestForm.timeframe)
        await msg.answer(get_text(lang, "bt_timeframe"), reply_markup=back_kb(lang))
    except:
        await msg.answer(get_text(lang, "error_date"))

@router.message(BacktestForm.timeframe)
async def bt_tf(msg: Message, state: FSMContext):
    await state.update_data(timeframe=msg.text.upper())
    await state.set_state(BacktestForm.commission)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "bt_commission"), reply_markup=back_kb(lang))

@router.message(BacktestForm.commission)
async def bt_comm(msg: Message, state: FSMContext):
    try:
        await state.update_data(commission=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.spread)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "bt_spread"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.spread)
async def bt_spread(msg: Message, state: FSMContext):
    try:
        await state.update_data(spread=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.asset)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.asset)
async def bt_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(BacktestForm.direction)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang))

@router.callback_query(F.data.in_(["LONG","SHORT"]))
async def bt_dir(call: CallbackQuery, state: FSMContext):
    await state.update_data(direction=call.data)
    await state.set_state(BacktestForm.entry_price)
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(BacktestForm.entry_price)
async def bt_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.sl_price)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "enter_sl"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.sl_price)
async def bt_sl(msg: Message, state: FSMContext):
    try:
        await state.update_data(sl_price=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.tp_price)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "enter_tp"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.tp_price)
async def bt_tp(msg: Message, state: FSMContext):
    try:
        await state.update_data(tp_price=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.exit_price)
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "enter_exit_price_bt"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.exit_price)
async def bt_exit(msg: Message, state: FSMContext):
    try:
        exit_p = float(msg.text.replace(",","."))
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
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(f"📊 P&L: ${pnl:.2f} ({r:.2f}R)\n\n{get_text(lang, 'enter_entry_time')}", reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id) or "ru"
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.entry_time)
async def bt_etime(msg: Message, state: FSMContext):
    await state.update_data(entry_time=msg.text)
    await state.set_state(BacktestForm.exit_time)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "enter_exit_time"), reply_markup=back_kb(lang))

@router.message(BacktestForm.exit_time)
async def bt_xtime(msg: Message, state: FSMContext):
    await state.update_data(exit_time=msg.text)
    await state.set_state(BacktestForm.setup)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "enter_setup"), reply_markup=back_kb(lang))

@router.message(BacktestForm.setup)
async def bt_setup(msg: Message, state: FSMContext):
    await state.update_data(setup=msg.text)
    await state.set_state(BacktestForm.trigger)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "enter_trigger"), reply_markup=back_kb(lang))

@router.message(BacktestForm.trigger)
async def bt_trigger(msg: Message, state: FSMContext):
    await state.update_data(trigger=msg.text)
    await state.set_state(BacktestForm.signal_quality)
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer(get_text(lang, "enter_quality"), reply_markup=quality_kb())

@router.callback_query(F.data.startswith("q_"))
async def bt_quality(call: CallbackQuery, state: FSMContext):
    q = int(call.data.split("_")[1])
    await state.update_data(signal_quality=q)
    await state.set_state(BacktestForm.link_chart)
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text(get_text(lang, "enter_link_bt"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(BacktestForm.link_chart)
async def bt_link(msg: Message, state: FSMContext):
    link = msg.text if msg.text != "0" else "-"
    data = await state.get_data()
    bt_data = {
        'user_id': msg.from_user.id,
        'period_start': data['period_start'], 'period_end': data['period_end'],
        'timeframe': data['timeframe'], 'commission': data['commission'], 'spread': data['spread'],
        'asset': data['asset'], 'direction': data['direction'], 'entry_price': data['entry_price'],
        'exit_price': data['exit_price'], 'sl_price': data['sl_price'], 'tp_price': data['tp_price'],
        'pnl_usd': data['pnl_usd'], 'pnl_r': data['pnl_r'], 'signal_quality': data['signal_quality'],
        'setup': data['setup'], 'trigger': data['trigger'], 'link_chart': link,
        'entry_time': data['entry_time'], 'exit_time': data['exit_time']
    }
    save_backtest(bt_data)
    await state.clear()
    lang = get_user_lang(msg.from_user.id) or "ru"
    await msg.answer("✅ Бэктест сохранён!", reply_markup=main_menu_kb(lang, "backtest"))

# ========== СТАТИСТИКА И ЭКСПОРТ ==========
@router.callback_query(F.data == "stats_all")
async def stats_all(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid) or "ru"
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        df = get_trades(uid, mode="real")
        txt = get_real_stats_text(df, lang, "stats_header")
    else:
        df = get_backtests(uid)
        txt = get_backtest_stats_text(df, lang)
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=main_menu_kb(lang, mode))
    await call.answer()

@router.callback_query(F.data == "stats_day")
async def stats_day(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid) or "ru"
    today = datetime.now().strftime("%Y-%m-%d")
    df = get_trades(uid, mode="real", start_date=today)
    txt = get_real_stats_text(df, lang, "stats_today")
    mode = "real"
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=main_menu_kb(lang, mode))
    await call.answer()

@router.callback_query(F.data == "stats_week")
async def stats_week(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid) or "ru"
    week = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    df = get_trades(uid, mode="real", start_date=week)
    txt = get_real_stats_text(df, lang, "stats_week")
    mode = "real"
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=main_menu_kb(lang, mode))
    await call.answer()

@router.callback_query(F.data == "stats_month")
async def stats_month(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid) or "ru"
    month = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    df = get_trades(uid, mode="real", start_date=month)
    txt = get_real_stats_text(df, lang, "stats_month")
    mode = "real"
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=main_menu_kb(lang, mode))
    await call.answer()

@router.callback_query(F.data == "get_excel")
async def get_excel(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        df = get_trades(uid, mode="real")
        if df.empty:
            await call.answer(get_text(get_user_lang(uid) or "ru", "no_data"), show_alert=True)
            return
        fname = export_real_to_excel(df, uid)
    else:
        df = get_backtests(uid)
        if df.empty:
            await call.answer(get_text(get_user_lang(uid) or "ru", "no_data"), show_alert=True)
            return
        fname = export_backtest_to_excel(df, uid)
    await call.message.answer_document(FSInputFile(fname), caption="📊 Ваш отчёт")
    os.remove(fname)
    await call.answer()

@router.callback_query(F.data == "clear_confirm")
async def clear_confirm(call: CallbackQuery, state: FSMContext):
    lang = get_user_lang(call.from_user.id) or "ru"
    await call.message.edit_text("⚠️ Удалить все сделки?", reply_markup=confirm_kb(lang))
    await call.answer()

@router.callback_query(F.data == "clear_yes")
async def clear_yes(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        clear_trades(uid, mode="real")
    else:
        clear_backtests(uid)
    lang = get_user_lang(uid) or "ru"
    await call.message.edit_text(get_text(lang, "cleared"), reply_markup=main_menu_kb(lang, mode))
    await call.answer()

# ========== ЗАПУСК ==========
async def main():
    global bot
    init_dbs()
    bot = Bot(token=BOT_TOKEN)
    await bot.set_my_commands([
        BotCommand(command="start", description="Start / Запуск"),
        BotCommand(command="stats", description="All statistics / Вся статистика"),
        BotCommand(command="day", description="Today / Сегодня"),
        BotCommand(command="week", description="Week / Неделя"),
        BotCommand(command="month", description="Month / Месяц"),
        BotCommand(command="get", description="Excel report / Excel отчёт"),
        BotCommand(command="clear", description="Clear journal / Очистить журнал"),
    ])
    dp = Dispatcher()
    dp.include_router(router)
    print("✅ Бот запущен! Режимы: реальная торговля и бэктест, Excel отдельно для каждого.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
