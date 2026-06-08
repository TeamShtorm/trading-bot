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
    "en": {
        "select_mode": "🎛 **Select mode:**",
        "mode_real": "📊 Real Trading",
        "mode_backtest": "🔄 Backtest",
        "add_trade": "➕ Add Trade",
        "stats": "📊 Statistics",
        "excel": "📎 Excel",
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
        "language_set": "✅ Language set: English",
        "language_set_ru": "✅ Language set: Russian",
        "support_info": "📞 **Support**\n\nIf you have any questions, problems, or suggestions — write to our support channel:\n\n👉 **@TJsupport_bot**\n\n**How to get help:**\n1. Open the channel\n2. Read the pinned post\n3. Write your question in the **comments** under the pinned post\n4. Describe the problem in detail, attach a screenshot if possible\n\nWe will reply as soon as possible!\n\n📌 **For quick processing, include your Telegram ID** (you can get it from @userinfobot).",
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
        "bt_stats_header": "📊 Backtest statistics",
        "bt_total_trades": "📋 Total trades: {total}",
        "bt_wins": "✅ Winning: {wins}",
        "bt_losses": "❌ Losing: {losses}",
        "bt_winrate": "🎯 Win rate: {wr:.1f}%",
        "bt_avg_r": "📊 Avg R: {avg_r:.2f}",
        "bt_total_r": "💰 Total R: {total_r:.2f}",
        "bt_quality": "⭐ Signal quality: {q:.1f}/5",
        "select_asset": "💰 **Select asset:**",
        "select_period": "📅 **Select period:**",
        "period_day": "📆 Day",
        "period_week": "📅 Week",
        "period_month": "📊 Month",
        "emotion_calm": "😊 Calm",
        "emotion_fear": "😨 Fear",
        "emotion_greed": "😈 Greed",
        "emotion_tilt": "🤬 Tilt",
        "emotion_confidence": "😌 Confidence",
        "trade_detail": "📋 **Trade #{id}**\n\nAsset: {asset}\nDirection: {direction}\nEntry: ${entry}\nExit: ${exit}\nVolume: {volume}\nP&L: ${pnl}\nOutcome: {result}\nDate: {date}\nLinks:\n{links}",
        "recent_trades": "🕒 **Recent trades** (last {count}):",
        "sort_newest": "📅 Newest first",
        "sort_oldest": "📅 Oldest first",
        "bt_trade_detail": "📋 **Backtest #{id}**\n\nAsset: {asset}\nDirection: {direction}\nEntry: ${entry}\nExit: ${exit}\nP&L: ${pnl} ({r:.2f}R)\nQuality: {quality}/5\nSetup: {setup}\nTrigger: {trigger}\nDate: {date}",
        "excel_ready": "📊 Your report",
        "change_lang": "🌐 Change language"
    },
    "ru": {
        "select_mode": "🎛 **Выберите режим работы:**",
        "mode_real": "📊 Реальная торговля",
        "mode_backtest": "🔄 Бэктест",
        "add_trade": "➕ Сделка",
        "stats": "📊 Статистика",
        "excel": "📎 Excel",
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
        "language_set": "✅ Язык установлен: английский",
        "language_set_ru": "✅ Язык установлен: русский",
        "support_info": "📞 **Поддержка**\n\nЕсли у вас есть вопросы, проблемы или предложения — пишите в наш канал поддержки:\n\n👉 **@TJsupport_bot**\n\n**Как получить помощь:**\n1. Откройте канал\n2. Прочитайте закреплённое сообщение\n3. Напишите свой вопрос в **комментариях** под закреплённым постом\n4. Опишите проблему подробно, приложите скриншот\n\nМы ответим как можно скорее!\n\n📌 **Для быстрой обработки укажите ваш Telegram ID** (можно узнать у @userinfobot).",
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
        "bt_stats_header": "📊 Статистика бэктеста",
        "bt_total_trades": "📋 Всего сделок: {total}",
        "bt_wins": "✅ Прибыльных: {wins}",
        "bt_losses": "❌ Убыточных: {losses}",
        "bt_winrate": "🎯 Винрейт: {wr:.1f}%",
        "bt_avg_r": "📊 Средний R: {avg_r:.2f}",
        "bt_total_r": "💰 Суммарный R: {total_r:.2f}",
        "bt_quality": "⭐ Качество сигнала: {q:.1f}/5",
        "select_asset": "💰 **Выберите актив:**",
        "select_period": "📅 **Выберите период:**",
        "period_day": "📆 День",
        "period_week": "📅 Неделя",
        "period_month": "📊 Месяц",
        "emotion_calm": "😊 Спокойствие",
        "emotion_fear": "😨 Страх",
        "emotion_greed": "😈 Жадность",
        "emotion_tilt": "🤬 Тильт",
        "emotion_confidence": "😌 Уверенность",
        "trade_detail": "📋 **Сделка #{id}**\n\nАктив: {asset}\nНаправление: {direction}\nВход: ${entry}\nВыход: ${exit}\nОбъём: {volume}\nP&L: ${pnl}\nИсход: {result}\nДата: {date}\nСсылки:\n{links}",
        "recent_trades": "🕒 **Недавние сделки** (последние {count}):",
        "sort_newest": "📅 Сначала новые",
        "sort_oldest": "📅 Сначала старые",
        "bt_trade_detail": "📋 **Бэктест #{id}**\n\nАктив: {asset}\nНаправление: {direction}\nВход: ${entry}\nВыход: ${exit}\nP&L: ${pnl} ({r:.2f}R)\nКачество: {quality}/5\nСетап: {setup}\nТриггер: {trigger}\nДата: {date}",
        "excel_ready": "📊 Ваш отчёт",
        "change_lang": "🌐 Сменить язык"
    }
}

def get_text(lang, key, **kwargs):
    t = TEXTS.get(lang, TEXTS["en"]).get(key, key)
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
            result TEXT, comment TEXT, trade_date TEXT, links TEXT, emotion TEXT
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
    conn.execute("CREATE TABLE IF NOT EXISTS users (user_id INTEGER PRIMARY KEY, lang TEXT DEFAULT 'en')")
    conn.commit()
    conn.close()

def get_user_lang(user_id):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT lang FROM users WHERE user_id = ?", (user_id,))
    r = cur.fetchone()
    conn.close()
    return r[0] if r else "en"

def set_user_lang(user_id, lang):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("INSERT OR REPLACE INTO users (user_id, lang) VALUES (?, ?)", (user_id, lang))
    conn.commit()
    conn.close()

def save_trade(user_id, mode, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links, emotion=""):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        INSERT INTO trades (user_id, mode, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links, emotion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, mode, asset, direction, entry_price, exit_price, volume, pnl, result, comment, trade_date, links, emotion))
    conn.commit()
    conn.close()

def get_trades(user_id, mode=None, start_date=None, asset=None, sort_by_date="DESC", limit=None):
    conn = sqlite3.connect(DB_NAME)
    q = "SELECT * FROM trades WHERE user_id = ?"
    p = [user_id]
    if mode:
        q += " AND mode = ?"
        p.append(mode)
    if start_date:
        q += " AND trade_date >= ?"
        p.append(start_date)
    if asset:
        q += " AND asset = ?"
        p.append(asset)
    q += f" ORDER BY trade_date {sort_by_date}"
    if limit:
        q += f" LIMIT {limit}"
    df = pd.read_sql_query(q, conn, params=p)
    conn.close()
    return df

def get_trade_by_id(trade_id, user_id):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT * FROM trades WHERE id = ? AND user_id = ?", (trade_id, user_id))
    row = cur.fetchone()
    conn.close()
    if row:
        cols = ['id', 'user_id', 'mode', 'asset', 'direction', 'entry_price', 'exit_price', 'volume', 'pnl', 'result', 'comment', 'trade_date', 'links', 'emotion']
        return dict(zip(cols, row))
    return None

def get_all_assets(user_id, mode):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT asset FROM trades WHERE user_id = ? AND mode = ?", (user_id, mode))
    assets = [row[0] for row in cur.fetchall()]
    conn.close()
    return assets

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

def get_backtests(user_id, limit=None):
    conn = sqlite3.connect(BT_DB_NAME)
    q = "SELECT * FROM backtests WHERE user_id = ? ORDER BY period_start DESC"
    p = [user_id]
    if limit:
        q += f" LIMIT {limit}"
    df = pd.read_sql_query(q, conn, params=p)
    conn.close()
    return df

def get_backtest_by_id(bt_id, user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT * FROM backtests WHERE id = ? AND user_id = ?", (bt_id, user_id))
    row = cur.fetchone()
    conn.close()
    if row:
        cols = ['id', 'user_id', 'period_start', 'period_end', 'timeframe', 'commission', 'spread', 'asset', 'direction',
                'entry_price', 'exit_price', 'sl_price', 'tp_price', 'pnl_usd', 'pnl_r', 'signal_quality', 'setup', 'trigger', 'link_chart', 'entry_time', 'exit_time']
        return dict(zip(cols, row))
    return None

def get_backtest_assets(user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT asset FROM backtests WHERE user_id = ?", (user_id,))
    assets = [row[0] for row in cur.fetchall()]
    conn.close()
    return assets

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
    df_exp.columns = ['📅 Date', '🪙 Asset', '📈 Direction', '💰 Entry', '💰 Exit', '📊 Volume', '💵 P&L', '🎯 Outcome', '📝 Comment', '🔗 Links', '😊 Emotion']
    df_exp['📈 Direction'] = df_exp['📈 Direction'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_exp['🎯 Outcome'] = df_exp['🎯 Outcome'].replace({'TAKE': '✅ Take', 'STOP': '❌ Stop'})
    df_exp = df_exp.sort_values('📅 Date', ascending=False)
    fname = f"real_{user_id}.xlsx"
    with pd.ExcelWriter(fname, engine='openpyxl') as w:
        df_exp.to_excel(w, sheet_name='Real Trading', index=False)
        ws = w.sheets['Real Trading']
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
    df_exp.columns = ['📅 Start', '📅 End', '⏱ Timeframe', '🪙 Asset', '📈 Direction',
                      '💰 Entry', '💰 Exit', '💵 P&L', '📊 P&L (R)', '⭐ Quality', '🎯 Setup', '⚡ Trigger']
    df_exp['📈 Direction'] = df_exp['📈 Direction'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_exp = df_exp.sort_values('📅 Start', ascending=False)
    fname = f"backtest_{user_id}.xlsx"
    with pd.ExcelWriter(fname, engine='openpyxl') as w:
        df_exp.to_excel(w, sheet_name='Backtest', index=False)
        ws = w.sheets['Backtest']
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
    emotions = df['emotion'].value_counts().to_dict()
    emotion_text = "\n".join([f"{e}: {c}" for e, c in emotions.items()]) if emotions else get_text(lang, "no_data")
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
        f"{get_text(lang, 'pf', pf=pf)}\n\n"
        f"😊 **Emotions:**\n{emotion_text}"
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

def mode_menu_kb(lang, mode):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "add_trade"), callback_data="add_trade")],
        [InlineKeyboardButton(text=get_text(lang, "stats"), callback_data="stats_menu")],
        [InlineKeyboardButton(text=get_text(lang, "excel"), callback_data="get_excel")],
        [InlineKeyboardButton(text=get_text(lang, "clear"), callback_data="clear_confirm")],
        [InlineKeyboardButton(text=get_text(lang, "settings"), callback_data="settings_menu"),
         InlineKeyboardButton(text=get_text(lang, "support"), callback_data="support")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def stats_menu_kb(lang):
    buttons = [
        [InlineKeyboardButton(text=get_text(lang, "stats_all"), callback_data="stats_all")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_asset"), callback_data="stats_by_asset")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_date"), callback_data="stats_by_date")],
        [InlineKeyboardButton(text=get_text(lang, "stats_recent"), callback_data="stats_recent")],
        [InlineKeyboardButton(text=get_text(lang, "stats_sort"), callback_data="stats_sort")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons + [[InlineKeyboardButton(text=get_text(lang, "stats_back"), callback_data="back_to_mode_menu")]])

def stats_menu_kb_with_emotions(lang):
    buttons = [
        [InlineKeyboardButton(text=get_text(lang, "stats_all"), callback_data="stats_all")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_asset"), callback_data="stats_by_asset")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_date"), callback_data="stats_by_date")],
        [InlineKeyboardButton(text=get_text(lang, "stats_by_emotion"), callback_data="stats_by_emotion")],
        [InlineKeyboardButton(text=get_text(lang, "stats_recent"), callback_data="stats_recent")],
        [InlineKeyboardButton(text=get_text(lang, "stats_sort"), callback_data="stats_sort")],
        [InlineKeyboardButton(text=get_text(lang, "stats_back"), callback_data="back_to_mode_menu")]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def asset_kb(assets, lang):
    buttons = [[InlineKeyboardButton(text=a, callback_data=f"asset_{a}")] for a in assets]
    buttons.append([InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def period_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "period_day"), callback_data="period_day")],
        [InlineKeyboardButton(text=get_text(lang, "period_week"), callback_data="period_week")],
        [InlineKeyboardButton(text=get_text(lang, "period_month"), callback_data="period_month")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ])

def emotion_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "emotion_calm"), callback_data="emotion_calm")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_fear"), callback_data="emotion_fear")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_greed"), callback_data="emotion_greed")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_tilt"), callback_data="emotion_tilt")],
        [InlineKeyboardButton(text=get_text(lang, "emotion_confidence"), callback_data="emotion_confidence")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ])

def sort_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "sort_newest"), callback_data="sort_newest")],
        [InlineKeyboardButton(text=get_text(lang, "sort_oldest"), callback_data="sort_oldest")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")]
    ])

def recent_trades_kb(trades, lang, mode):
    buttons = []
    for _, row in trades.iterrows():
        pnl = row['pnl'] if mode == "real" else row['pnl_usd']
        emoji = "✅" if pnl > 0 else "❌"
        buttons.append([InlineKeyboardButton(text=f"{emoji} #{row['id']} {row['asset']} | ${pnl:.0f}", callback_data=f"trade_{row['id']}")])
    buttons.append([InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def back_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def confirm_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "confirm_clear"), callback_data="clear_yes")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_to_mode_menu")]
    ])

def settings_kb(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text(lang, "change_lang"), callback_data="change_lang")],
        [InlineKeyboardButton(text=get_text(lang, "support"), callback_data="support")],
        [InlineKeyboardButton(text=get_text(lang, "back"), callback_data="back_mode")]
    ])

def lang_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇬🇧 English", callback_data="lang_en")],
        [InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru")]
    ])

def quality_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1", callback_data="q_1"), InlineKeyboardButton(text="2", callback_data="q_2"),
         InlineKeyboardButton(text="3", callback_data="q_3"), InlineKeyboardButton(text="4", callback_data="q_4"),
         InlineKeyboardButton(text="5", callback_data="q_5")]
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
router = Router()
bot = None

# ========== СТАРТ И ОСНОВНЫЕ ОБРАБОТЧИКИ ==========
@router.message(CommandStart())
async def start_cmd(msg: Message, state: FSMContext):
    await state.clear()
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    await msg.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))

@router.message(Command("settings"))
async def settings_cmd(msg: Message, state: FSMContext):
    await state.clear()
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    await msg.answer(get_text(lang, "settings_menu"), parse_mode="Markdown", reply_markup=settings_kb(lang))

@router.callback_query(F.data == "settings_menu")
async def settings(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "settings_menu"), parse_mode="Markdown", reply_markup=settings_kb(lang))
    await call.answer()

@router.callback_query(F.data == "change_lang")
async def change_lang(call: CallbackQuery):
    await call.message.edit_text(get_text("en", "select_language"), reply_markup=lang_kb())
    await call.answer()

@router.callback_query(F.data.startswith("lang_"))
async def set_lang(call: CallbackQuery, state: FSMContext):
    lang = call.data.split("_")[1]
    set_user_lang(call.from_user.id, lang)
    await call.message.delete()
    await call.message.answer(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

@router.callback_query(F.data == "support")
async def support(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    text = get_text(lang, "support_info")
    await call.message.edit_text(
        text,
        parse_mode="Markdown",
        reply_markup=mode_kb(lang)
    )
    await call.answer()

@router.callback_query(F.data == "back_mode")
async def back_mode(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

@router.callback_query(F.data == "back_to_mode_menu")
async def back_to_mode_menu(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    mode = (await state.get_data()).get('mode', 'real')
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_menu_kb(lang, mode))
    await call.answer()

# ========== ВЫБОР РЕЖИМА ==========
@router.callback_query(F.data.in_(["mode_real", "mode_backtest"]))
async def choose_mode(call: CallbackQuery, state: FSMContext):
    mode = call.data.split("_")[1]
    await state.update_data(mode=mode)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_menu_kb(lang, mode))
    await call.answer()

# ========== СТАТИСТИКА МЕНЮ ==========
@router.callback_query(F.data == "stats_menu")
async def stats_menu(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    mode = data.get('mode', 'real')
    lang = get_user_lang(call.from_user.id)
    if mode == "real":
        await call.message.edit_text(get_text(lang, "stats_menu"), parse_mode="Markdown", reply_markup=stats_menu_kb_with_emotions(lang))
    else:
        await call.message.edit_text(get_text(lang, "stats_menu"), parse_mode="Markdown", reply_markup=stats_menu_kb(lang))
    await call.answer()

# ========== ВСЯ СТАТИСТИКА ==========
@router.callback_query(F.data == "stats_all")
async def stats_all(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    sort_order = data.get('sort_order', 'DESC')
    if mode == "real":
        df = get_trades(uid, mode="real", sort_by_date=sort_order)
        txt = get_real_stats_text(df, lang, "stats_header")
    else:
        df = get_backtests(uid)
        txt = get_backtest_stats_text(df, lang)
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=stats_menu_kb_with_emotions(lang) if mode == "real" else stats_menu_kb(lang))
    await call.answer()

# ========== ПО АКТИВАМ ==========
@router.callback_query(F.data == "stats_by_asset")
async def stats_by_asset(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        assets = get_all_assets(uid, mode="real")
    else:
        assets = get_backtest_assets(uid)
    if not assets:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    await call.message.edit_text(get_text(lang, "select_asset"), parse_mode="Markdown", reply_markup=asset_kb(assets, lang))
    await call.answer()

@router.callback_query(F.data.startswith("asset_"))
async def show_asset_stats(call: CallbackQuery, state: FSMContext):
    asset = call.data.split("_")[1]
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    sort_order = data.get('sort_order', 'DESC')
    if mode == "real":
        df = get_trades(uid, mode="real", asset=asset, sort_by_date=sort_order)
        txt = get_real_stats_text(df, lang, f"📊 Statistics for {asset}")
    else:
        df = get_backtests(uid)
        df = df[df['asset'] == asset]
        txt = get_backtest_stats_text(df, lang)
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=asset_kb(get_all_assets(uid, mode) if mode == "real" else get_backtest_assets(uid), lang))
    await call.answer()

# ========== ПО ДАТЕ ==========
@router.callback_query(F.data == "stats_by_date")
async def stats_by_date(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_period"), parse_mode="Markdown", reply_markup=period_kb(lang))
    await call.answer()

@router.callback_query(F.data.in_(["period_day", "period_week", "period_month"]))
async def show_period_stats(call: CallbackQuery, state: FSMContext):
    period = call.data.split("_")[1]
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    sort_order = data.get('sort_order', 'DESC')
    if period == "day":
        start_date = datetime.now().strftime("%Y-%m-%d")
        title = get_text(lang, "stats_today")
    elif period == "week":
        start_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
        title = get_text(lang, "stats_week")
    else:
        start_date = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
        title = get_text(lang, "stats_month")
    if mode == "real":
        df = get_trades(uid, mode="real", start_date=start_date, sort_by_date=sort_order)
        txt = get_real_stats_text(df, lang, title)
    else:
        df = get_backtests(uid)
        txt = get_backtest_stats_text(df, lang)
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=period_kb(lang))
    await call.answer()

# ========== ПО ЭМОЦИЯМ ==========
@router.callback_query(F.data == "stats_by_emotion")
async def stats_by_emotion(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_period") + " (Emotions)", parse_mode="Markdown", reply_markup=emotion_kb(lang))
    await call.answer()

@router.callback_query(F.data.startswith("emotion_"))
async def show_emotion_stats(call: CallbackQuery, state: FSMContext):
    emotion_map = {
        "emotion_calm": "😊 Calm",
        "emotion_fear": "😨 Fear",
        "emotion_greed": "😈 Greed",
        "emotion_tilt": "🤬 Tilt",
        "emotion_confidence": "😌 Confidence"
    }
    emotion = emotion_map.get(call.data, "😊 Calm")
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    sort_order = data.get('sort_order', 'DESC')
    df = get_trades(uid, mode="real", sort_by_date=sort_order)
    df = df[df['emotion'] == emotion]
    txt = get_real_stats_text(df, lang, f"😊 {emotion}")
    await call.message.edit_text(txt, parse_mode="Markdown", reply_markup=emotion_kb(lang))
    await call.answer()

# ========== НЕДАВНИЕ СДЕЛКИ ==========
@router.callback_query(F.data == "stats_recent")
async def stats_recent(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        df = get_trades(uid, mode="real", sort_by_date="DESC", limit=10)
        if df.empty:
            await call.answer(get_text(lang, "no_data"), show_alert=True)
            return
        await call.message.edit_text(get_text(lang, "recent_trades", count=len(df)), parse_mode="Markdown", reply_markup=recent_trades_kb(df, lang, mode))
    else:
        df = get_backtests(uid, limit=10)
        if df.empty:
            await call.answer(get_text(lang, "no_data"), show_alert=True)
            return
        buttons = []
        for _, row in df.iterrows():
            pnl = row['pnl_usd']
            emoji = "✅" if pnl > 0 else "❌"
            buttons.append([InlineKeyboardButton(text=f"{emoji} #{row['id']} {row['asset']} | ${pnl:.0f}", callback_data=f"bt_{row['id']}")])
        buttons.append([InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_menu")])
        await call.message.edit_text(get_text(lang, "recent_trades", count=len(df)), parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("trade_"))
async def show_trade_detail(call: CallbackQuery):
    trade_id = int(call.data.split("_")[1])
    uid = call.from_user.id
    lang = get_user_lang(uid)
    trade = get_trade_by_id(trade_id, uid)
    if not trade:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    links = trade.get('links', '') or '-'
    text = get_text(lang, "trade_detail",
        id=trade['id'], asset=trade['asset'], direction="🟢 LONG" if trade['direction'] == "LONG" else "🔴 SHORT",
        entry=trade['entry_price'], exit=trade['exit_price'], volume=trade['volume'],
        pnl=trade['pnl'], result="✅ TAKE" if trade['result'] == "TAKE" else "❌ STOP",
        date=trade['trade_date'], links=links
    )
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_recent")]]))
    await call.answer()

@router.callback_query(F.data.startswith("bt_"))
async def show_backtest_detail(call: CallbackQuery):
    bt_id = int(call.data.split("_")[1])
    uid = call.from_user.id
    lang = get_user_lang(uid)
    bt = get_backtest_by_id(bt_id, uid)
    if not bt:
        await call.answer(get_text(lang, "no_data"), show_alert=True)
        return
    text = get_text(lang, "bt_trade_detail",
        id=bt['id'], asset=bt['asset'], direction="🟢 LONG" if bt['direction'] == "LONG" else "🔴 SHORT",
        entry=bt['entry_price'], exit=bt['exit_price'], pnl=bt['pnl_usd'], r=bt['pnl_r'],
        quality=bt['signal_quality'], setup=bt['setup'], trigger=bt['trigger'], date=bt['period_start']
    )
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text(lang, "back"), callback_data="stats_recent")]]))
    await call.answer()

# ========== СОРТИРОВКА ==========
@router.callback_query(F.data == "stats_sort")
async def stats_sort(call: CallbackQuery):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "stats_sort"), parse_mode="Markdown", reply_markup=sort_kb(lang))
    await call.answer()

@router.callback_query(F.data.in_(["sort_newest", "sort_oldest"]))
async def set_sort_order(call: CallbackQuery, state: FSMContext):
    sort_order = "DESC" if call.data == "sort_newest" else "ASC"
    await state.update_data(sort_order=sort_order)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "select_mode"), parse_mode="Markdown", reply_markup=mode_kb(lang))
    await call.answer()

# ========== ДОБАВЛЕНИЕ СДЕЛКИ (РЕАЛЬНАЯ) ==========
@router.callback_query(F.data == "add_trade")
async def add_trade_start(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        await state.set_state(TradeForm.asset)
        lang = get_user_lang(call.from_user.id)
        await call.message.edit_text(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))
    else:
        await state.set_state(BacktestForm.period_start)
        lang = get_user_lang(call.from_user.id)
        await call.message.edit_text(get_text(lang, "bt_period_start"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.asset)
async def real_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(TradeForm.direction)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang))

@router.callback_query(F.data.in_(["LONG","SHORT"]))
async def real_dir(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    mode = data.get('mode', 'real')
    await state.update_data(direction=call.data)
    if mode == "real":
        await state.set_state(TradeForm.entry_price)
        lang = get_user_lang(call.from_user.id)
        await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    else:
        await state.set_state(BacktestForm.entry_price)
        lang = get_user_lang(call.from_user.id)
        await call.message.edit_text(get_text(lang, "enter_entry_price"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.entry_price)
async def real_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",",".")))
        await state.set_state(TradeForm.exit_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_exit_price"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(TradeForm.exit_price)
async def real_exit(msg: Message, state: FSMContext):
    try:
        await state.update_data(exit_price=float(msg.text.replace(",",".")))
        await state.set_state(TradeForm.volume)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_volume"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(TradeForm.volume)
async def real_vol(msg: Message, state: FSMContext):
    try:
        await state.update_data(volume=float(msg.text.replace(",",".")))
        await state.set_state(TradeForm.result)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "choose_result"), reply_markup=result_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.callback_query(F.data.in_(["TAKE","STOP"]))
async def real_res(call: CallbackQuery, state: FSMContext):
    await state.update_data(result=call.data)
    await state.set_state(TradeForm.comment)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_comment"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.comment)
async def real_comment(msg: Message, state: FSMContext):
    com = msg.text.strip()
    if com == "-":
        com = ""
    await state.update_data(comment=com)
    await state.set_state(TradeForm.add_link)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "add_link_question"), reply_markup=yesno_kb(lang))

@router.callback_query(F.data == "yes")
async def add_link_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(TradeForm.link_url)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_link"), reply_markup=back_kb(lang))
    await call.answer()

@router.callback_query(F.data == "no")
async def skip_links(call: CallbackQuery, state: FSMContext):
    await state.update_data(links="")
    await state.set_state(TradeForm.trade_date)
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text(get_text(lang, "enter_date"), reply_markup=back_kb(lang))
    await call.answer()

@router.message(TradeForm.link_url)
async def get_link(msg: Message, state: FSMContext):
    url = msg.text
    await state.update_data(link_url=url)
    await state.set_state(TradeForm.link_tf)
    lang = get_user_lang(msg.from_user.id)
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
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "link_saved"), reply_markup=yesno_kb(lang))

@router.message(TradeForm.trade_date)
async def real_date(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id)
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
    await state.update_data(trade_date=tdate, pnl=pnl)
    await state.set_state(TradeForm.emotion)
    await msg.answer(get_text(lang, "enter_emotion"), reply_markup=emotion_kb(lang))

@router.callback_query(F.data.startswith("emotion_"))
async def real_emotion(call: CallbackQuery, state: FSMContext):
    emotion_map = {
        "emotion_calm": "😊 Calm",
        "emotion_fear": "😨 Fear",
        "emotion_greed": "😈 Greed",
        "emotion_tilt": "🤬 Tilt",
        "emotion_confidence": "😌 Confidence"
    }
    emotion = emotion_map.get(call.data, "😊 Calm")
    data = await state.get_data()
    save_trade(
        call.from_user.id, "real", data['asset'], data['direction'],
        data['entry_price'], data['exit_price'], data['volume'], data['pnl'],
        data['result'], data['comment'], data['trade_date'], data.get('links', ''), emotion
    )
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    mode = data.get('mode', 'real')
    await call.message.edit_text(get_text(lang, "trade_saved"), parse_mode="Markdown", reply_markup=mode_menu_kb(lang, mode))
    await call.answer()

# ========== БЭКТЕСТ (ДОБАВЛЕНИЕ) ==========
@router.message(BacktestForm.period_start)
async def bt_start(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id)
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_start=d)
        await state.set_state(BacktestForm.period_end)
        await msg.answer(get_text(lang, "bt_period_end"), reply_markup=back_kb(lang))
    except:
        await msg.answer(get_text(lang, "error_date"))

@router.message(BacktestForm.period_end)
async def bt_end(msg: Message, state: FSMContext):
    lang = get_user_lang(msg.from_user.id)
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
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "bt_commission"), reply_markup=back_kb(lang))

@router.message(BacktestForm.commission)
async def bt_comm(msg: Message, state: FSMContext):
    try:
        await state.update_data(commission=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.spread)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "bt_spread"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.spread)
async def bt_spread(msg: Message, state: FSMContext):
    try:
        await state.update_data(spread=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.asset)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_asset"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.asset)
async def bt_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(BacktestForm.direction)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "choose_direction"), reply_markup=direction_kb(lang))

@router.message(BacktestForm.entry_price)
async def bt_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.sl_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_sl"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.sl_price)
async def bt_sl(msg: Message, state: FSMContext):
    try:
        await state.update_data(sl_price=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.tp_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_tp"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.tp_price)
async def bt_tp(msg: Message, state: FSMContext):
    try:
        await state.update_data(tp_price=float(msg.text.replace(",",".")))
        await state.set_state(BacktestForm.exit_price)
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "enter_exit_price_bt"), reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
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
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(f"📊 P&L: ${pnl:.2f} ({r:.2f}R)\n\n{get_text(lang, 'enter_entry_time')}", reply_markup=back_kb(lang))
    except:
        lang = get_user_lang(msg.from_user.id)
        await msg.answer(get_text(lang, "error_number"))

@router.message(BacktestForm.entry_time)
async def bt_etime(msg: Message, state: FSMContext):
    await state.update_data(entry_time=msg.text)
    await state.set_state(BacktestForm.exit_time)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_exit_time"), reply_markup=back_kb(lang))

@router.message(BacktestForm.exit_time)
async def bt_xtime(msg: Message, state: FSMContext):
    await state.update_data(exit_time=msg.text)
    await state.set_state(BacktestForm.setup)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_setup"), reply_markup=back_kb(lang))

@router.message(BacktestForm.setup)
async def bt_setup(msg: Message, state: FSMContext):
    await state.update_data(setup=msg.text)
    await state.set_state(BacktestForm.trigger)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_trigger"), reply_markup=back_kb(lang))

@router.message(BacktestForm.trigger)
async def bt_trigger(msg: Message, state: FSMContext):
    await state.update_data(trigger=msg.text)
    await state.set_state(BacktestForm.signal_quality)
    lang = get_user_lang(msg.from_user.id)
    await msg.answer(get_text(lang, "enter_quality"), reply_markup=quality_kb())

@router.callback_query(F.data.startswith("q_"))
async def bt_quality(call: CallbackQuery, state: FSMContext):
    q = int(call.data.split("_")[1])
    await state.update_data(signal_quality=q)
    await state.set_state(BacktestForm.link_chart)
    lang = get_user_lang(call.from_user.id)
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
    lang = get_user_lang(msg.from_user.id)
    await msg.answer("✅ Backtest saved!", reply_markup=mode_menu_kb(lang, "backtest"))

# ========== EXCEL И ОЧИСТКА ==========
@router.message(Command("get"))
async def cmd_get(msg: Message, state: FSMContext):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    
    if mode == "real":
        df = get_trades(uid, mode="real")
        if df.empty:
            await msg.answer(get_text(lang, "no_data_add_trade"))
            return
        fname = export_real_to_excel(df, uid)
    else:
        df = get_backtests(uid)
        if df.empty:
            await msg.answer(get_text(lang, "no_data_add_trade"))
            return
        fname = export_backtest_to_excel(df, uid)
    
    await msg.answer_document(FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    os.remove(fname)

@router.callback_query(F.data == "get_excel")
async def get_excel(call: CallbackQuery, state: FSMContext):
    uid = call.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    
    if mode == "real":
        df = get_trades(uid, mode="real")
        if df.empty:
            await call.answer(get_text(lang, "no_data_add_trade"), show_alert=True)
            return
        fname = export_real_to_excel(df, uid)
    else:
        df = get_backtests(uid)
        if df.empty:
            await call.answer(get_text(lang, "no_data_add_trade"), show_alert=True)
            return
        fname = export_backtest_to_excel(df, uid)
    
    await call.message.answer_document(FSInputFile(fname), caption=get_text(lang, "excel_ready"))
    os.remove(fname)
    await call.answer()

@router.callback_query(F.data == "clear_confirm")
async def clear_confirm(call: CallbackQuery, state: FSMContext):
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text("⚠️ Delete all trades?", reply_markup=confirm_kb(lang))
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
    lang = get_user_lang(uid)
    await call.message.edit_text(get_text(lang, "cleared"), reply_markup=mode_menu_kb(lang, mode))
    await call.answer()

# ========== КОМАНДЫ ==========
@router.message(Command("stats"))
async def cmd_stats(msg: Message, state: FSMContext):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    sort_order = data.get('sort_order', 'DESC')
    if mode == "real":
        df = get_trades(uid, mode="real", sort_by_date=sort_order)
        txt = get_real_stats_text(df, lang, "stats_header")
    else:
        df = get_backtests(uid)
        txt = get_backtest_stats_text(df, lang)
    await msg.answer(txt, parse_mode="Markdown")

@router.message(Command("day"))
async def cmd_day(msg: Message, state: FSMContext):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    today = datetime.now().strftime("%Y-%m-%d")
    df = get_trades(uid, mode="real", start_date=today)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    txt = get_real_stats_text(df, lang, "stats_today")
    await msg.answer(txt, parse_mode="Markdown")

@router.message(Command("week"))
async def cmd_week(msg: Message, state: FSMContext):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    week_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    df = get_trades(uid, mode="real", start_date=week_ago)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    txt = get_real_stats_text(df, lang, "stats_week")
    await msg.answer(txt, parse_mode="Markdown")

@router.message(Command("month"))
async def cmd_month(msg: Message, state: FSMContext):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    month_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    df = get_trades(uid, mode="real", start_date=month_ago)
    if df.empty:
        await msg.answer(get_text(lang, "no_data_add_trade"))
        return
    txt = get_real_stats_text(df, lang, "stats_month")
    await msg.answer(txt, parse_mode="Markdown")

@router.message(Command("clear"))
async def cmd_clear(msg: Message, state: FSMContext):
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    data = await state.get_data()
    mode = data.get('mode', 'real')
    if mode == "real":
        clear_trades(uid, mode="real")
    else:
        clear_backtests(uid)
    await msg.answer(get_text(lang, "cleared"))

# ========== ЗАПУСК ==========
async def main():
    global bot
    init_dbs()
    bot = Bot(token=BOT_TOKEN)
    await bot.set_my_commands([
        BotCommand(command="start", description="Start / Запуск"),
        BotCommand(command="settings", description="Settings / Настройки"),
        BotCommand(command="stats", description="All statistics / Вся статистика"),
        BotCommand(command="day", description="Today / Сегодня"),
        BotCommand(command="week", description="Week / Неделя"),
        BotCommand(command="month", description="Month / Месяц"),
        BotCommand(command="get", description="Excel report / Excel отчёт"),
        BotCommand(command="clear", description="Clear journal / Очистить журнал"),
    ])
    dp = Dispatcher()
    dp.include_router(router)
    print("✅ Бот запущен! Поддержка добавлена в главное меню.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
