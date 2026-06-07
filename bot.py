import asyncio
import sqlite3
import os
import re
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher, F, Router
from aiogram.types import Message, CallbackQuery, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton, BotCommand
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ========== КОНФИГ ==========
BOT_TOKEN = "8584035526:AAG8Q15ym8TONEAOH4_8_eQaXnsV4VhhIYs"
REAL_DB = "trades.db"
BACKTEST_DB = "backtests.db"

# ========== БАЗА ДАННЫХ (реальная торговля) ==========
def init_real_db():
    conn = sqlite3.connect(REAL_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            asset TEXT,
            direction TEXT,
            volume REAL,
            entry_price REAL,
            exit_price REAL,
            sl_price REAL,
            tp_price REAL,
            pnl_usd REAL,
            pnl_percent REAL,
            commission REAL,
            setup TEXT,
            trigger TEXT,
            emotion TEXT,
            discipline TEXT,
            lesson TEXT,
            link_15m TEXT,
            link_1h TEXT,
            link_4h TEXT,
            link_1d TEXT,
            link_1w TEXT,
            link_1m TEXT,
            entry_time TEXT,
            exit_time TEXT,
            trade_date TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_real_trade(data):
    conn = sqlite3.connect(REAL_DB)
    conn.execute("""
        INSERT INTO trades (
            user_id, asset, direction, volume, entry_price, exit_price,
            sl_price, tp_price, pnl_usd, pnl_percent, commission,
            setup, trigger, emotion, discipline, lesson,
            link_15m, link_1h, link_4h, link_1d, link_1w, link_1m,
            entry_time, exit_time, trade_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data['user_id'], data['asset'], data['direction'], data['volume'],
        data['entry_price'], data['exit_price'], data['sl_price'], data['tp_price'],
        data['pnl_usd'], data['pnl_percent'], data['commission'],
        data['setup'], data['trigger'], data['emotion'], data['discipline'], data['lesson'],
        data['link_15m'], data['link_1h'], data['link_4h'], data['link_1d'],
        data['link_1w'], data['link_1m'], data['entry_time'], data['exit_time'], data['trade_date']
    ))
    conn.commit()
    conn.close()

def get_real_trades(user_id, start_date=None):
    conn = sqlite3.connect(REAL_DB)
    query = "SELECT * FROM trades WHERE user_id = ?"
    params = [user_id]
    if start_date:
        try:
            d = datetime.strptime(start_date, "%d.%m.%Y")
            sql_date = d.strftime("%Y-%m-%d")
            query += " AND substr(trade_date, 7, 4)||'-'||substr(trade_date, 4, 2)||'-'||substr(trade_date, 1, 2) >= ?"
            params.append(sql_date)
        except:
            pass
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_real_assets(user_id):
    conn = sqlite3.connect(REAL_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT asset FROM trades WHERE user_id = ?", (user_id,))
    assets = [row[0] for row in cursor.fetchall()]
    conn.close()
    return assets

def clear_real_trades(user_id):
    conn = sqlite3.connect(REAL_DB)
    conn.execute("DELETE FROM trades WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

# ========== БАЗА ДАННЫХ (бэктест) ==========
def init_backtest_db():
    conn = sqlite3.connect(BACKTEST_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS backtests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            period_start TEXT,
            period_end TEXT,
            timeframe TEXT,
            commission_percent REAL,
            spread REAL,
            asset TEXT,
            direction TEXT,
            entry_price REAL,
            exit_price REAL,
            sl_price REAL,
            tp_price REAL,
            pnl_usd REAL,
            pnl_percent REAL,
            pnl_r REAL,
            mae REAL,
            mfe REAL,
            signal_quality INTEGER,
            skipped BOOLEAN,
            setup TEXT,
            trigger TEXT,
            link_chart TEXT,
            entry_time TEXT,
            exit_time TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_backtest(data):
    conn = sqlite3.connect(BACKTEST_DB)
    conn.execute("""
        INSERT INTO backtests (
            user_id, period_start, period_end, timeframe, commission_percent, spread,
            asset, direction, entry_price, exit_price, sl_price, tp_price,
            pnl_usd, pnl_percent, pnl_r, mae, mfe,
            signal_quality, skipped, setup, trigger, link_chart, entry_time, exit_time
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data['user_id'], data['period_start'], data['period_end'], data['timeframe'],
        data['commission_percent'], data['spread'], data['asset'], data['direction'],
        data['entry_price'], data['exit_price'], data['sl_price'], data['tp_price'],
        data['pnl_usd'], data['pnl_percent'], data['pnl_r'], data['mae'], data['mfe'],
        data['signal_quality'], data['skipped'], data['setup'], data['trigger'],
        data['link_chart'], data['entry_time'], data['exit_time']
    ))
    conn.commit()
    conn.close()

def get_backtests(user_id, start_date=None):
    conn = sqlite3.connect(BACKTEST_DB)
    query = "SELECT * FROM backtests WHERE user_id = ?"
    params = [user_id]
    if start_date:
        query += " AND entry_time >= ?"
        params.append(start_date)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def clear_backtests(user_id):
    conn = sqlite3.connect(BACKTEST_DB)
    conn.execute("DELETE FROM backtests WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

# ========== КРАСИВЫЙ ЭКСПОРТ В EXCEL (реальная торговля) ==========
def export_real_to_excel(df, user_id):
    df_export = df.copy()
    df_export = df_export[[
        'trade_date', 'entry_time', 'asset', 'direction', 'volume',
        'entry_price', 'exit_price', 'sl_price', 'tp_price',
        'pnl_usd', 'pnl_percent', 'commission',
        'setup', 'trigger', 'emotion', 'discipline', 'lesson'
    ]]
    
    df_export.columns = [
        '📅 Дата', '⏰ Время', '🪙 Монета', '📈 Направление', '📊 Объём',
        '💰 Вход', '💰 Выход', '🛑 SL', '🎯 TP',
        '💵 P&L ($)', '📊 P&L (%)', '💸 Комиссия',
        '🎯 Сетап', '⚡ Триггер', '😊 Эмоции', '📋 Дисциплина', '📝 Вывод'
    ]
    
    df_export['📈 Направление'] = df_export['📈 Направление'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_export = df_export.sort_values('📅 Дата', ascending=False)
    
    fname = f"real_journal_{user_id}.xlsx"
    
    with pd.ExcelWriter(fname, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='Реальная торговля', index=False)
        worksheet = writer.sheets['Реальная торговля']
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2b6cb0", end_color="2b6cb0", fill_type="solid")
        for col in range(1, len(df_export.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
        
        green_fill = PatternFill(start_color="c6f7d0", end_color="c6f7d0", fill_type="solid")
        red_fill = PatternFill(start_color="fecaca", end_color="fecaca", fill_type="solid")
        
        for row in range(2, len(df_export) + 2):
            pnl_cell = worksheet.cell(row=row, column=11)
            if pnl_cell.value and pnl_cell.value > 0:
                for col in range(1, len(df_export.columns) + 1):
                    worksheet.cell(row=row, column=col).fill = green_fill
            elif pnl_cell.value and pnl_cell.value < 0:
                for col in range(1, len(df_export.columns) + 1):
                    worksheet.cell(row=row, column=col).fill = red_fill
        
        for col in range(1, len(df_export.columns) + 1):
            max_len = 0
            col_letter = get_column_letter(col)
            for row in range(1, len(df_export) + 2):
                val = worksheet.cell(row=row, column=col).value
                if val:
                    max_len = max(max_len, len(str(val)))
            worksheet.column_dimensions[col_letter].width = min(max_len + 2, 30)
        worksheet.freeze_panes = 'A2'
    
    return fname

def export_backtest_to_excel(df, user_id):
    df_export = df.copy()
    df_export = df_export[[
        'period_start', 'period_end', 'timeframe', 'asset', 'direction',
        'entry_price', 'exit_price', 'sl_price', 'tp_price',
        'pnl_usd', 'pnl_r', 'mae', 'mfe',
        'signal_quality', 'skipped', 'setup', 'trigger'
    ]]
    
    df_export.columns = [
        '📅 Начало', '📅 Конец', '⏱ Таймфрейм', '🪙 Монета', '📈 Направление',
        '💰 Вход', '💰 Выход', '🛑 SL', '🎯 TP',
        '💵 P&L ($)', '📊 P&L (R)', '📉 MAE', '📈 MFE',
        '⭐ Качество', '❌ Пропуск', '🎯 Сетап', '⚡ Триггер'
    ]
    
    df_export['📈 Направление'] = df_export['📈 Направление'].replace({'LONG': '🟢 LONG', 'SHORT': '🔴 SHORT'})
    df_export['❌ Пропуск'] = df_export['❌ Пропуск'].replace({True: '✅ Пропустил', False: '❌ Вошёл'})
    df_export = df_export.sort_values('📅 Начало', ascending=False)
    
    fname = f"backtest_journal_{user_id}.xlsx"
    
    with pd.ExcelWriter(fname, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='Бэктест', index=False)
        worksheet = writer.sheets['Бэктест']
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2b6cb0", end_color="2b6cb0", fill_type="solid")
        for col in range(1, len(df_export.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
        
        for col in range(1, len(df_export.columns) + 1):
            max_len = 0
            col_letter = get_column_letter(col)
            for row in range(1, len(df_export) + 2):
                val = worksheet.cell(row=row, column=col).value
                if val:
                    max_len = max(max_len, len(str(val)))
            worksheet.column_dimensions[col_letter].width = min(max_len + 2, 30)
        worksheet.freeze_panes = 'A2'
    
    return fname

# ========== СТАТИСТИКА ==========
def calc_real_stats(df, user_id, name="Все сделки"):
    if df.empty:
        return "📭 Нет данных", None
    
    pnl = df['pnl_usd']
    total = len(pnl)
    takes = len(pnl[pnl > 0])
    stops = len(pnl[pnl < 0])
    wr = (takes / total * 100) if total > 0 else 0
    
    sp = pnl[pnl > 0].sum()
    sl = abs(pnl[pnl < 0].sum())
    pf = sp / sl if sl > 0 else sp
    
    mean = pnl.mean()
    best = pnl.max()
    worst = pnl.min()
    cum = pnl.cumsum()
    dd = (cum.cummax() - cum).max()
    
    emotions = df['emotion'].value_counts().to_dict()
    emotion_text = "\n".join([f"  {e}: {c}" for e, c in emotions.items()]) if emotions else "  нет данных"
    
    discipline_yes = len(df[df['discipline'] == '✅ Строго по плану'])
    discipline_no = len(df[df['discipline'] == '❌ Нарушил правила'])
    discipline_pct = (discipline_yes / total * 100) if total > 0 else 0
    
    plt.figure(figsize=(10, 5))
    plt.plot(cum, marker='o', color='blue', linewidth=2)
    plt.axhline(0, color='red', linestyle='--')
    plt.title(f"Кривая доходности — {name}")
    plt.grid(True)
    path = f"real_chart_{user_id}.png"
    plt.savefig(path, bbox_inches='tight')
    plt.close()
    
    text = (
        f"📊 **{name}**\n\n"
        f"📋 Сделок: {total} | ✅ {takes} | ❌ {stops}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"⚙️ Профит-фактор: {pf:.2f}\n"
        f"📉 Просадка: ${dd:.2f}\n"
        f"💰 Средняя: ${mean:.2f}\n"
        f"🏆 Лучший: +${best:.2f} | 💀 Худший: ${worst:.2f}\n\n"
        f"😊 Эмоции:\n{emotion_text}\n\n"
        f"📋 Дисциплина: по плану {discipline_yes} ({discipline_pct:.0f}%), нарушил {discipline_no}"
    )
    return text, path

def calc_backtest_stats(df, user_id):
    if df.empty:
        return "📭 Нет данных", None
    
    pnl = df['pnl_usd']
    total = len(pnl)
    takes = len(pnl[pnl > 0])
    stops = len(pnl[pnl < 0])
    wr = (takes / total * 100) if total > 0 else 0
    
    sp = pnl[pnl > 0].sum()
    sl = abs(pnl[pnl < 0].sum())
    pf = sp / sl if sl > 0 else sp
    
    r_vals = df['pnl_r']
    avg_r = r_vals.mean()
    total_r = r_vals.sum()
    avg_quality = df['signal_quality'].mean()
    
    cum = pnl.cumsum()
    dd = (cum.cummax() - cum).max()
    
    plt.figure(figsize=(10, 5))
    plt.plot(cum, marker='o', color='green', linewidth=2)
    plt.axhline(0, color='red', linestyle='--')
    plt.title("Кривая доходности бэктеста")
    plt.grid(True)
    path = f"backtest_chart_{user_id}.png"
    plt.savefig(path, bbox_inches='tight')
    plt.close()
    
    text = (
        f"📊 **Бэктест**\n\n"
        f"📋 Сделок: {total} | ✅ {takes} | ❌ {stops}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"⚙️ Профит-фактор: {pf:.2f}\n"
        f"📉 Просадка: ${dd:.2f}\n"
        f"💰 Средняя сделка: ${pnl.mean():.2f}\n"
        f"📊 Средний R: {avg_r:.2f} | Суммарный R: {total_r:.2f}\n"
        f"⭐ Качество сигнала: {avg_quality:.1f}/5"
    )
    return text, path

def del_chart(path):
    if path and os.path.exists(path):
        os.remove(path)

# ========== ВСПОМОГАТЕЛЬНЫЕ ==========
async def safe_delete(msg):
    try:
        await msg.delete()
    except:
        pass

async def delete_previous_messages(state, key):
    data = await state.get_data()
    msg_id = data.get(key)
    if msg_id:
        try:
            await bot.delete_message(chat_id=data.get('chat_id'), message_id=msg_id)
        except:
            pass

def parse_date(date_str):
    date_str = date_str.strip().lower()
    if date_str == "сегодня":
        return datetime.now().strftime("%d.%m.%Y")
    match = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", date_str)
    if match:
        d, m, y = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= d <= 31 and 1 <= m <= 12 and 1900 <= y <= 2100:
            return f"{d:02d}.{m:02d}.{y}"
    if date_str.isdigit() and len(date_str) == 8:
        d, m, y = int(date_str[0:2]), int(date_str[2:4]), int(date_str[4:8])
        if 1 <= d <= 31 and 1 <= m <= 12 and 1900 <= y <= 2100:
            return f"{d:02d}.{m:02d}.{y}"
    return None

def parse_time(time_str):
    match = re.match(r"(\d{1,2}):(\d{2})", time_str.strip())
    if match:
        h, m = int(match.group(1)), int(match.group(2))
        if 0 <= h <= 23 and 0 <= m <= 59:
            return f"{h:02d}:{m:02d}"
    return None

def calc_pnl(direction, entry, exit_price, volume, commission):
    if direction == "LONG":
        pnl = (exit_price - entry) * volume - commission
    else:
        pnl = (entry - exit_price) * volume - commission
    pnl_pct = (pnl / (entry * volume)) * 100 if entry * volume > 0 else 0
    return pnl, pnl_pct

def calc_r(direction, entry, sl, exit_price):
    risk = abs(entry - sl)
    if direction == "LONG":
        pnl = exit_price - entry
    else:
        pnl = entry - exit_price
    return pnl / risk if risk > 0 else 0

# ========== КНОПКИ ==========
def main_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Режим: Реальная торговля", callback_data="mode_real")],
        [InlineKeyboardButton(text="🔄 Режим: Бэктест", callback_data="mode_backtest")],
        [InlineKeyboardButton(text="📈 Статистика", callback_data="stats_main")],
        [InlineKeyboardButton(text="🧮 Калькулятор риска", callback_data="calc_risk")]
    ])

def real_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Новая сделка", callback_data="add_real")],
        [InlineKeyboardButton(text="📊 Статистика", callback_data="real_stats")],
        [InlineKeyboardButton(text="📎 Excel отчёт", callback_data="real_excel")],
        [InlineKeyboardButton(text="🗑 Очистить", callback_data="real_clear")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")]
    ])

def backtest_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔄 Новая сделка", callback_data="add_backtest")],
        [InlineKeyboardButton(text="📊 Статистика", callback_data="backtest_stats")],
        [InlineKeyboardButton(text="📎 Excel отчёт", callback_data="backtest_excel")],
        [InlineKeyboardButton(text="🗑 Очистить", callback_data="backtest_clear")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")]
    ])

def direction_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🟢 LONG", callback_data="LONG"),
         InlineKeyboardButton(text="🔴 SHORT", callback_data="SHORT")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_mode")]
    ])

def emotion_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="😊 Спокойствие", callback_data="emotion_calm"),
         InlineKeyboardButton(text="😨 Страх", callback_data="emotion_fear")],
        [InlineKeyboardButton(text="😈 Жадность", callback_data="emotion_greed"),
         InlineKeyboardButton(text="🤬 Тильт", callback_data="emotion_tilt")],
        [InlineKeyboardButton(text="😌 Уверенность", callback_data="emotion_confidence")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_mode")]
    ])

def discipline_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Строго по плану", callback_data="discipline_yes")],
        [InlineKeyboardButton(text="❌ Нарушил правила", callback_data="discipline_no")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_mode")]
    ])

def quality_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⭐1", callback_data="q_1"), InlineKeyboardButton(text="⭐⭐2", callback_data="q_2"),
         InlineKeyboardButton(text="⭐⭐⭐3", callback_data="q_3"), InlineKeyboardButton(text="⭐⭐⭐⭐4", callback_data="q_4"),
         InlineKeyboardButton(text="⭐⭐⭐⭐⭐5", callback_data="q_5")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_mode")]
    ])

def skip_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Вошёл бы", callback_data="skip_no")],
        [InlineKeyboardButton(text="❌ Пропустил бы", callback_data="skip_yes")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_mode")]
    ])

def cancel():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_mode")]
    ])

def confirm():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚠️ ДА, УДАЛИТЬ ВСЁ", callback_data="confirm_clear")],
        [InlineKeyboardButton(text="🔙 Отмена", callback_data="back_to_mode")]
    ])

def stats_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📈 Вся статистика", callback_data="stats_all")],
        [InlineKeyboardButton(text="💰 По монетам", callback_data="stats_assets")],
        [InlineKeyboardButton(text="📅 По дате", callback_data="stats_time")],
        [InlineKeyboardButton(text="😊 По эмоциям", callback_data="stats_emotions")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")]
    ])

def time_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📆 За день", callback_data="time_day")],
        [InlineKeyboardButton(text="📅 За неделю", callback_data="time_week")],
        [InlineKeyboardButton(text="📊 За месяц", callback_data="time_month")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="stats_main")]
    ])

# ========== FSM ==========
class RealTradeForm(StatesGroup):
    asset = State()
    direction = State()
    volume = State()
    entry_price = State()
    sl_price = State()
    tp_price = State()
    exit_price = State()
    commission = State()
    entry_time = State()
    exit_time = State()
    trade_date = State()
    setup = State()
    trigger = State()
    emotion = State()
    discipline = State()
    lesson = State()
    link15 = State()
    link1h = State()
    link4h = State()
    link1d = State()
    link1w = State()
    link1m = State()

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
    mae = State()
    mfe = State()
    setup = State()
    trigger = State()
    signal_quality = State()
    skipped = State()
    link_chart = State()

class RiskForm(StatesGroup):
    depo = State()
    percent = State()
    entry = State()
    stop = State()

# ========== РОУТЕРЫ ==========
router = Router()
bot = None

# ========== КОМАНДЫ ==========
@router.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await state.clear()
    await safe_delete(msg)
    sent = await msg.answer(
        "📊 **Trading Journal**\n\n"
        "Выберите режим работы:\n"
        "• 📊 Реальная торговля — запись реальных сделок\n"
        "• 🔄 Бэктест — тестирование стратегий\n\n"
        "📌 Команды:\n"
        "/start - Главное меню\n"
        "/stats - Статистика\n"
        "/cancel - Отменить действие",
        parse_mode="Markdown",
        reply_markup=main_menu()
    )
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)

@router.message(Command("stats"))
async def cmd_stats(msg: Message, state: FSMContext):
    await state.clear()
    await safe_delete(msg)
    sent = await msg.answer("📊 Выберите тип статистики:", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)

@router.message(Command("cancel"))
async def cmd_cancel(msg: Message, state: FSMContext):
    await state.clear()
    await safe_delete(msg)
    sent = await msg.answer("✅ Действие отменено.", reply_markup=main_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)

# ========== ГЛАВНОЕ МЕНЮ ==========
@router.callback_query(F.data == "back_to_main")
async def back_to_main(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    sent = await call.message.answer("🏠 Главное меню:", reply_markup=main_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.callback_query(F.data == "back_to_mode")
async def back_to_mode(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    sent = await call.message.answer("🏠 Главное меню:", reply_markup=main_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.callback_query(F.data == "mode_real")
async def mode_real(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    await state.update_data(mode="real")
    sent = await call.message.answer(
        "📊 **Режим: Реальная торговля**\n\n"
        "Записывайте свои сделки с полной аналитикой:\n"
        "• Стоп-лосс и тейк-профит\n"
        "• Эмоции и дисциплина\n"
        "• P&L в $ и %\n"
        "• Ссылки на графики\n\n"
        "👇 Выберите действие:",
        parse_mode="Markdown",
        reply_markup=real_menu()
    )
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.callback_query(F.data == "mode_backtest")
async def mode_backtest(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    await state.update_data(mode="backtest")
    sent = await call.message.answer(
        "🔄 **Режим: Бэктест**\n\n"
        "Тестируйте стратегии на истории:\n"
        "• Период и таймфрейм\n"
        "• Спред и комиссии\n"
        "• R (риск-метрика)\n"
        "• MAE/MFE\n"
        "• Качество сигнала 1-5\n\n"
        "👇 Выберите действие:",
        parse_mode="Markdown",
        reply_markup=backtest_menu()
    )
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

# ========== РЕАЛЬНАЯ ТОРГОВЛЯ ==========
@router.callback_query(F.data == "add_real")
async def add_real(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    await state.update_data(mode="real")
    await state.set_state(RealTradeForm.asset)
    sent = await call.message.answer("📝 Введите тикер (BTC, ETH, TON, AAPL):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.message(RealTradeForm.asset)
async def real_asset(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(asset=msg.text.upper())
    await state.set_state(RealTradeForm.direction)
    sent = await msg.answer("📈 Направление:", reply_markup=direction_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data.in_(["LONG", "SHORT"]))
async def real_direction(call: CallbackQuery, state: FSMContext):
    await safe_delete(call.message)
    await state.update_data(direction=call.data)
    mode = (await state.get_data()).get('mode', 'real')
    if mode == "real":
        await state.set_state(RealTradeForm.volume)
        sent = await call.message.answer("📊 Объём позиции (размер в $):", reply_markup=cancel())
    else:
        await state.set_state(BacktestForm.entry_price)
        sent = await call.message.answer("💰 Цена входа:", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)
    await call.answer()

@router.message(RealTradeForm.volume)
async def real_volume(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(volume=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.entry_price)
        sent = await msg.answer("💰 Цена входа:", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.entry_price)
async def real_entry(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(entry_price=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.sl_price)
        sent = await msg.answer("🛑 Стоп-Лосс (0 если нет):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.sl_price)
async def real_sl(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(sl_price=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.tp_price)
        sent = await msg.answer("🎯 Тейк-Профит (0 если нет):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.tp_price)
async def real_tp(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(tp_price=float(msg.text.replace(",", ".")))
        await state.set_state(RealTradeForm.exit_price)
        sent = await msg.answer("💰 Цена выхода (реальная):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.exit_price)
async def real_exit(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        exit_price = float(msg.text.replace(",", "."))
        data = await state.get_data()
        pnl, pnl_pct = calc_pnl(data['direction'], data['entry_price'], exit_price, data['volume'], 0)
        await state.update_data(exit_price=exit_price, pnl_usd=pnl, pnl_percent=pnl_pct)
        await state.set_state(RealTradeForm.commission)
        sent = await msg.answer(f"💸 Комиссия биржи в $ (0 если нет):\n(Предварительный P&L: ${pnl:.2f})", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.commission)
async def real_commission(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        commission = float(msg.text.replace(",", "."))
        data = await state.get_data()
        pnl, pnl_pct = calc_pnl(data['direction'], data['entry_price'], data['exit_price'], data['volume'], commission)
        await state.update_data(commission=commission, pnl_usd=pnl, pnl_percent=pnl_pct)
        await state.set_state(RealTradeForm.entry_time)
        sent = await msg.answer(f"💰 Итоговый P&L: ${pnl:.2f} ({pnl_pct:.1f}%)\n\n⏰ Время входа (ЧЧ:ММ):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.entry_time)
async def real_entry_time(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_time(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ЧЧ:ММ", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(entry_time=parsed)
    await state.set_state(RealTradeForm.exit_time)
    sent = await msg.answer("⏰ Время выхода (ЧЧ:ММ):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.exit_time)
async def real_exit_time(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_time(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ЧЧ:ММ", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(exit_time=parsed)
    await state.set_state(RealTradeForm.trade_date)
    sent = await msg.answer("📅 Дата сделки (ДД.ММ.ГГГГ) или 'сегодня':", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.trade_date)
async def real_trade_date(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_date(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ДД.ММ.ГГГГ или 'сегодня'", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(trade_date=parsed)
    await state.set_state(RealTradeForm.setup)
    sent = await msg.answer("🎯 Сетап (стратегия/паттерн):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.setup)
async def real_setup(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(setup=msg.text)
    await state.set_state(RealTradeForm.trigger)
    sent = await msg.answer("⚡ Триггер (сигнал индикатора, объём, новость):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.trigger)
async def real_trigger(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(trigger=msg.text)
    await state.set_state(RealTradeForm.emotion)
    sent = await msg.answer("😊 Эмоции во время сделки:", reply_markup=emotion_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data.startswith("emotion_"))
async def real_emotion(call: CallbackQuery, state: FSMContext):
    await safe_delete(call.message)
    emotions = {"emotion_calm": "😊 Спокойствие", "emotion_fear": "😨 Страх",
                "emotion_greed": "😈 Жадность", "emotion_tilt": "🤬 Тильт",
                "emotion_confidence": "😌 Уверенность"}
    await state.update_data(emotion=emotions.get(call.data, "😊 Спокойствие"))
    await state.set_state(RealTradeForm.discipline)
    sent = await call.message.answer("📋 Строго по плану?", reply_markup=discipline_menu())
    await state.update_data(last_bot_msg=sent.message_id)
    await call.answer()

@router.callback_query(F.data.startswith("discipline_"))
async def real_discipline(call: CallbackQuery, state: FSMContext):
    await safe_delete(call.message)
    discipline = "✅ Строго по плану" if call.data == "discipline_yes" else "❌ Нарушил правила"
    await state.update_data(discipline=discipline)
    await state.set_state(RealTradeForm.lesson)
    sent = await call.message.answer("📝 Вывод из сделки (что поняли/усвоили):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)
    await call.answer()

@router.message(RealTradeForm.lesson)
async def real_lesson(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(lesson=msg.text)
    await state.set_state(RealTradeForm.link15)
    sent = await msg.answer("🔗 Ссылка график 15м (0 нет):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

async def save_real_link(msg, state, field, next_state, next_text):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    val = msg.text.strip()
    await state.update_data({field: "-" if val == "0" else val})
    await state.set_state(next_state)
    sent = await msg.answer(next_text, reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(RealTradeForm.link15)
async def real_l15(msg: Message, state: FSMContext):
    await save_real_link(msg, state, "link_15m", RealTradeForm.link1h, "🔗 Ссылка 1ч (0 нет):")
@router.message(RealTradeForm.link1h)
async def real_l1h(msg: Message, state: FSMContext):
    await save_real_link(msg, state, "link_1h", RealTradeForm.link4h, "🔗 Ссылка 4ч (0 нет):")
@router.message(RealTradeForm.link4h)
async def real_l4h(msg: Message, state: FSMContext):
    await save_real_link(msg, state, "link_4h", RealTradeForm.link1d, "🔗 Ссылка 1д (0 нет):")
@router.message(RealTradeForm.link1d)
async def real_l1d(msg: Message, state: FSMContext):
    await save_real_link(msg, state, "link_1d", RealTradeForm.link1w, "🔗 Ссылка 1н (0 нет):")
@router.message(RealTradeForm.link1w)
async def real_l1w(msg: Message, state: FSMContext):
    await save_real_link(msg, state, "link_1w", RealTradeForm.link1m, "🔗 Ссылка 1м (0 нет):")

@router.message(RealTradeForm.link1m)
async def real_l1m(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    val = msg.text.strip()
    data = await state.get_data()
    
    trade_data = {
        'user_id': msg.from_user.id,
        'asset': data['asset'], 'direction': data['direction'], 'volume': data['volume'],
        'entry_price': data['entry_price'], 'exit_price': data['exit_price'],
        'sl_price': data.get('sl_price', 0), 'tp_price': data.get('tp_price', 0),
        'pnl_usd': data['pnl_usd'], 'pnl_percent': data['pnl_percent'], 'commission': data['commission'],
        'setup': data['setup'], 'trigger': data['trigger'], 'emotion': data['emotion'],
        'discipline': data['discipline'], 'lesson': data['lesson'],
        'link_15m': data.get('link_15m', '-'), 'link_1h': data.get('link_1h', '-'),
        'link_4h': data.get('link_4h', '-'), 'link_1d': data.get('link_1d', '-'),
        'link_1w': data.get('link_1w', '-'), 'link_1m': "-" if val == "0" else val,
        'entry_time': data['entry_time'], 'exit_time': data['exit_time'], 'trade_date': data['trade_date']
    }
    save_real_trade(trade_data)
    await state.clear()
    emoji = "✅" if trade_data['pnl_usd'] > 0 else "❌"
    sent = await msg.answer(f"{emoji} **Сделка сохранена!**\n💰 P&L: ${trade_data['pnl_usd']:.2f} ({trade_data['pnl_percent']:.1f}%)", parse_mode="Markdown", reply_markup=real_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)

# ========== БЭКТЕСТ ==========
@router.callback_query(F.data == "add_backtest")
async def add_backtest(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    await state.update_data(mode="backtest")
    await state.set_state(BacktestForm.period_start)
    sent = await call.message.answer("📅 Начало периода (ДД.ММ.ГГГГ):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.message(BacktestForm.period_start)
async def bt_period_start(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_date(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ДД.ММ.ГГГГ", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(period_start=parsed)
    await state.set_state(BacktestForm.period_end)
    sent = await msg.answer("📅 Конец периода (ДД.ММ.ГГГГ):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.period_end)
async def bt_period_end(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_date(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ДД.ММ.ГГГГ", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(period_end=parsed)
    await state.set_state(BacktestForm.timeframe)
    sent = await msg.answer("⏱ Таймфрейм (M5, H1, H4, D1, W1):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.timeframe)
async def bt_timeframe(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(timeframe=msg.text.upper())
    await state.set_state(BacktestForm.commission)
    sent = await msg.answer("💸 Комиссия в % (0.1):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.commission)
async def bt_commission(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(commission_percent=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.spread)
        sent = await msg.answer("📊 Спред в пунктах (1.5):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.spread)
async def bt_spread(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(spread=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.asset)
        sent = await msg.answer("🪙 Тикер (BTC, ETH):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.asset)
async def bt_asset(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(asset=msg.text.upper())
    await state.set_state(BacktestForm.direction)
    sent = await msg.answer("📈 Направление:", reply_markup=direction_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.entry_price)
async def bt_entry(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(entry_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.sl_price)
        sent = await msg.answer("🛑 Стоп-Лосс:", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.sl_price)
async def bt_sl(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(sl_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.tp_price)
        sent = await msg.answer("🎯 Тейк-Профит (0 если нет):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.tp_price)
async def bt_tp(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(tp_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.exit_price)
        sent = await msg.answer("💰 Цена выхода:", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.exit_price)
async def bt_exit(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        exit_price = float(msg.text.replace(",", "."))
        data = await state.get_data()
        direction = data['direction']
        entry = data['entry_price']
        sl = data['sl_price']
        spread = data['spread']
        commission_pct = data['commission_percent'] / 100
        
        if direction == "LONG":
            pnl_usd = (exit_price - entry) - commission_pct * entry - spread
            risk = abs(entry - sl)
        else:
            pnl_usd = (entry - exit_price) - commission_pct * entry - spread
            risk = abs(sl - entry)
        pnl_r = pnl_usd / risk if risk > 0 else 0
        pnl_pct = (pnl_usd / entry) * 100
        
        await state.update_data(exit_price=exit_price, pnl_usd=pnl_usd, pnl_percent=pnl_pct, pnl_r=pnl_r)
        await state.set_state(BacktestForm.entry_time)
        sent = await msg.answer(f"📊 P&L: ${pnl_usd:.2f} ({pnl_r:.2f}R)\n\n⏰ Время входа (ЧЧ:ММ):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.entry_time)
async def bt_entry_time(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_time(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ЧЧ:ММ", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(entry_time=parsed)
    await state.set_state(BacktestForm.exit_time)
    sent = await msg.answer("⏰ Время выхода (ЧЧ:ММ):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.exit_time)
async def bt_exit_time(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    parsed = parse_time(msg.text)
    if not parsed:
        sent = await msg.answer("❌ Формат ЧЧ:ММ", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await state.update_data(exit_time=parsed)
    await state.set_state(BacktestForm.mae)
    sent = await msg.answer("📉 MAE (макс. просадка в сделке в $):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.mae)
async def bt_mae(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(mae=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.mfe)
        sent = await msg.answer("📈 MFE (макс. прибыль в сделке в $):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.mfe)
async def bt_mfe(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(mfe=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.setup)
        sent = await msg.answer("🎯 Сетап (стратегия/паттерн):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Введите число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.setup)
async def bt_setup(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(setup=msg.text)
    await state.set_state(BacktestForm.trigger)
    sent = await msg.answer("⚡ Триггер:", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)

@router.message(BacktestForm.trigger)
async def bt_trigger(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    await state.update_data(trigger=msg.text)
    await state.set_state(BacktestForm.signal_quality)
    sent = await msg.answer("⭐ Качество сигнала (1-5):", reply_markup=quality_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data.startswith("q_"))
async def bt_quality(call: CallbackQuery, state: FSMContext):
    await safe_delete(call.message)
    await state.update_data(signal_quality=int(call.data.split("_")[1]))
    await state.set_state(BacktestForm.skipped)
    sent = await call.message.answer("❓ Пропустили бы сделку в реале?", reply_markup=skip_menu())
    await state.update_data(last_bot_msg=sent.message_id)
    await call.answer()

@router.callback_query(F.data.startswith("skip_"))
async def bt_skipped(call: CallbackQuery, state: FSMContext):
    await safe_delete(call.message)
    skipped = call.data == "skip_yes"
    await state.update_data(skipped=skipped)
    await state.set_state(BacktestForm.link_chart)
    sent = await call.message.answer("🔗 Ссылка на скриншот (0 если нет):", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id)
    await call.answer()

@router.message(BacktestForm.link_chart)
async def bt_link(call: CallbackQuery, state: FSMContext):
    await safe_delete(call.message)
    await delete_previous_messages(state, "last_bot_msg")
    val = call.message.text.strip()
    data = await state.get_data()
    
    backtest_data = {
        'user_id': call.from_user.id,
        'period_start': data['period_start'], 'period_end': data['period_end'],
        'timeframe': data['timeframe'], 'commission_percent': data['commission_percent'],
        'spread': data['spread'], 'asset': data['asset'], 'direction': data['direction'],
        'entry_price': data['entry_price'], 'exit_price': data['exit_price'],
        'sl_price': data['sl_price'], 'tp_price': data.get('tp_price', 0),
        'pnl_usd': data['pnl_usd'], 'pnl_percent': data['pnl_percent'], 'pnl_r': data['pnl_r'],
        'mae': data['mae'], 'mfe': data['mfe'], 'signal_quality': data['signal_quality'],
        'skipped': data['skipped'], 'setup': data['setup'], 'trigger': data['trigger'],
        'link_chart': "-" if val == "0" else val, 'entry_time': data['entry_time'], 'exit_time': data['exit_time']
    }
    save_backtest(backtest_data)
    await state.clear()
    emoji = "✅" if backtest_data['pnl_usd'] > 0 else "❌"
    sent = await call.message.answer(f"{emoji} **Сделка бэктеста сохранена!**\n💰 P&L: ${backtest_data['pnl_usd']:.2f} ({backtest_data['pnl_r']:.2f}R)", parse_mode="Markdown", reply_markup=backtest_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

# ========== СТАТИСТИКА (общая) ==========
@router.callback_query(F.data == "stats_main")
async def stats_main(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    sent = await call.message.answer("📊 Выберите тип статистики:", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.callback_query(F.data == "stats_all")
async def stats_all(call: CallbackQuery, state: FSMContext):
    await call.answer()
    df = get_real_trades(call.from_user.id)
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет данных реальной торговли.", reply_markup=stats_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await safe_delete(call.message)
    text, path = calc_real_stats(df, call.from_user.id)
    if path:
        sent = await call.message.answer_photo(photo=FSInputFile(path), caption=text, parse_mode="Markdown", reply_markup=stats_menu())
        del_chart(path)
    else:
        sent = await call.message.answer(text, parse_mode="Markdown", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "stats_assets")
async def stats_assets(call: CallbackQuery, state: FSMContext):
    await call.answer()
    assets = get_real_assets(call.from_user.id)
    if not assets:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет активов.", reply_markup=stats_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    btns = [[InlineKeyboardButton(text=f"🪙 {a}", callback_data=f"asset_{a}")] for a in assets]
    btns.append([InlineKeyboardButton(text="🔙 Назад", callback_data="stats_main")])
    await safe_delete(call.message)
    sent = await call.message.answer("💰 Выберите актив:", reply_markup=InlineKeyboardMarkup(inline_keyboard=btns))
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data.startswith("asset_"))
async def stats_asset_detail(call: CallbackQuery, state: FSMContext):
    await call.answer()
    asset = call.data.split("_")[1]
    df = get_real_trades(call.from_user.id)
    df = df[df['asset'] == asset]
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer(f"📭 Нет данных по {asset}", reply_markup=stats_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await safe_delete(call.message)
    text, path = calc_real_stats(df, call.from_user.id, asset)
    if path:
        sent = await call.message.answer_photo(photo=FSInputFile(path), caption=text, parse_mode="Markdown", reply_markup=stats_menu())
        del_chart(path)
    else:
        sent = await call.message.answer(text, parse_mode="Markdown", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "stats_time")
async def stats_time(call: CallbackQuery, state: FSMContext):
    await call.answer()
    await safe_delete(call.message)
    sent = await call.message.answer("📅 Выберите период:", reply_markup=time_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data.in_(["time_day", "time_week", "time_month"]))
async def stats_time_period(call: CallbackQuery, state: FSMContext):
    await call.answer()
    days = {"time_day": 1, "time_week": 7, "time_month": 30}
    start = (datetime.now() - timedelta(days=days[call.data])).strftime("%d.%m.%Y")
    df = get_real_trades(call.from_user.id, start_date=start)
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет сделок за период.", reply_markup=stats_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await safe_delete(call.message)
    names = {"time_day": "За день", "time_week": "За неделю", "time_month": "За месяц"}
    text, path = calc_real_stats(df, call.from_user.id, names[call.data])
    if path:
        sent = await call.message.answer_photo(photo=FSInputFile(path), caption=text, parse_mode="Markdown", reply_markup=stats_menu())
        del_chart(path)
    else:
        sent = await call.message.answer(text, parse_mode="Markdown", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "stats_emotions")
async def stats_emotions(call: CallbackQuery, state: FSMContext):
    await call.answer()
    df = get_real_trades(call.from_user.id)
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет данных.", reply_markup=stats_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    emotions = ['😊 Спокойствие', '😨 Страх', '😈 Жадность', '🤬 Тильт', '😌 Уверенность']
    text = "😊 **Статистика по эмоциям:**\n\n"
    for e in emotions:
        sub = df[df['emotion'] == e]
        if not sub.empty:
            total = len(sub)
            wins = len(sub[sub['pnl_usd'] > 0])
            wr = wins / total * 100
            avg = sub['pnl_usd'].mean()
            text += f"{e}: {total} сделок, винрейт {wr:.0f}%, средний ${avg:.0f}\n"
        else:
            text += f"{e}: 0 сделок\n"
    await safe_delete(call.message)
    sent = await call.message.answer(text, parse_mode="Markdown", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id)

# ========== EXCEL ==========
@router.callback_query(F.data == "real_excel")
async def real_excel(call: CallbackQuery, state: FSMContext):
    await call.answer()
    df = get_real_trades(call.from_user.id)
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет данных.", reply_markup=real_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    fname = export_real_to_excel(df, call.from_user.id)
    await safe_delete(call.message)
    await call.message.answer_document(document=FSInputFile(fname), caption="📎 Excel-отчёт (реальная торговля)")
    if os.path.exists(fname):
        os.remove(fname)
    sent = await call.message.answer("📊 Режим реальной торговли:", reply_markup=real_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "backtest_excel")
async def backtest_excel(call: CallbackQuery, state: FSMContext):
    await call.answer()
    df = get_backtests(call.from_user.id)
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет данных.", reply_markup=backtest_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    fname = export_backtest_to_excel(df, call.from_user.id)
    await safe_delete(call.message)
    await call.message.answer_document(document=FSInputFile(fname), caption="📎 Excel-отчёт (бэктест)")
    if os.path.exists(fname):
        os.remove(fname)
    sent = await call.message.answer("🔄 Режим бэктеста:", reply_markup=backtest_menu())
    await state.update_data(last_bot_msg=sent.message_id)

# ========== СТАТИСТИКА РЕЖИМОВ ==========
@router.callback_query(F.data == "real_stats")
async def real_stats_mode(call: CallbackQuery, state: FSMContext):
    await call.answer()
    await safe_delete(call.message)
    sent = await call.message.answer("📊 Статистика реальной торговли:", reply_markup=stats_menu())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "backtest_stats")
async def backtest_stats_mode(call: CallbackQuery, state: FSMContext):
    await call.answer()
    df = get_backtests(call.from_user.id)
    if df.empty:
        await safe_delete(call.message)
        sent = await call.message.answer("📭 Нет данных бэктеста.", reply_markup=backtest_menu())
        await state.update_data(last_bot_msg=sent.message_id)
        return
    await safe_delete(call.message)
    text, path = calc_backtest_stats(df, call.from_user.id)
    if path:
        sent = await call.message.answer_photo(photo=FSInputFile(path), caption=text, parse_mode="Markdown", reply_markup=backtest_menu())
        del_chart(path)
    else:
        sent = await call.message.answer(text, parse_mode="Markdown", reply_markup=backtest_menu())
    await state.update_data(last_bot_msg=sent.message_id)

# ========== ОЧИСТКА ==========
@router.callback_query(F.data == "real_clear")
async def real_clear_ask(call: CallbackQuery, state: FSMContext):
    await call.answer()
    await safe_delete(call.message)
    sent = await call.message.answer("⚠️ Удалить ВСЕ сделки реальной торговли?", reply_markup=confirm())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "backtest_clear")
async def backtest_clear_ask(call: CallbackQuery, state: FSMContext):
    await call.answer()
    await safe_delete(call.message)
    sent = await call.message.answer("⚠️ Удалить ВСЕ сделки бэктеста?", reply_markup=confirm())
    await state.update_data(last_bot_msg=sent.message_id)

@router.callback_query(F.data == "confirm_clear")
async def confirm_clear(call: CallbackQuery, state: FSMContext):
    await call.answer()
    mode = (await state.get_data()).get('mode', 'real')
    if mode == "real":
        clear_real_trades(call.from_user.id)
        await safe_delete(call.message)
        sent = await call.message.answer("🗑 Реальная торговля очищена!", reply_markup=real_menu())
    else:
        clear_backtests(call.from_user.id)
        await safe_delete(call.message)
        sent = await call.message.answer("🗑 Бэктест очищен!", reply_markup=backtest_menu())
    await state.update_data(last_bot_msg=sent.message_id)

# ========== КАЛЬКУЛЯТОР РИСКА ==========
@router.callback_query(F.data == "calc_risk")
async def calc_risk_start(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await safe_delete(call.message)
    await state.set_state(RiskForm.depo)
    sent = await call.message.answer("💰 Депозит в $:", reply_markup=cancel())
    await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    await call.answer()

@router.message(RiskForm.depo)
async def calc_depo(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(depo=float(msg.text.replace(",", ".")))
        await state.set_state(RiskForm.percent)
        sent = await msg.answer("📉 Риск % (1-2):", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RiskForm.percent)
async def calc_percent(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(percent=float(msg.text.replace(",", ".")))
        await state.set_state(RiskForm.entry)
        sent = await msg.answer("💰 Цена входа:", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RiskForm.entry)
async def calc_entry(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        await state.update_data(entry=float(msg.text.replace(",", ".")))
        await state.set_state(RiskForm.stop)
        sent = await msg.answer("🛑 Стоп-лосс:", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)
    except:
        sent = await msg.answer("❌ Ошибка! Число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

@router.message(RiskForm.stop)
async def calc_stop(msg: Message, state: FSMContext):
    await safe_delete(msg)
    await delete_previous_messages(state, "last_bot_msg")
    try:
        stop = float(msg.text.replace(",", "."))
        data = await state.get_data()
        loss_usd = data['depo'] * (data['percent'] / 100)
        diff = abs(data['entry'] - stop) / data['entry']
        if diff == 0:
            sent = await msg.answer("❌ Цена и стоп одинаковы!", reply_markup=cancel())
            await state.update_data(last_bot_msg=sent.message_id)
            return
        size = loss_usd / diff
        await state.clear()
        sent = await msg.answer(
            f"📊 **Результат:**\n\n"
            f"Убыток: ${loss_usd:.2f}\n"
            f"До стопа: {diff*100:.2f}%\n"
            f"Объём: ${size:.2f}",
            parse_mode="Markdown",
            reply_markup=main_menu()
        )
        await state.update_data(last_bot_msg=sent.message_id, chat_id=sent.chat.id)
    except:
        sent = await msg.answer("❌ Ошибка! Число.", reply_markup=cancel())
        await state.update_data(last_bot_msg=sent.message_id)

# ========== ЗАПУСК ==========
async def main():
    global bot
    init_real_db()
    init_backtest_db()
    bot = Bot(token=BOT_TOKEN)
    
    await bot.set_my_commands([
        BotCommand(command="start", description="🏠 Главное меню"),
        BotCommand(command="stats", description="📊 Статистика"),
        BotCommand(command="cancel", description="❌ Отменить действие"),
    ])
    
    dp = Dispatcher()
    dp.include_router(router)
    print("✅ Бот запущен! Режимы: реальная торговля + бэктест")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())