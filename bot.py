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
import matplotlib.pyplot as plt
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from aiohttp import web

# ========== КОНФИГ ==========
BOT_TOKEN = "8803530037:AAHVuMAb6gIzGXBKH8qbteZtFyttz6_hzh0"
DB_NAME = "trades.db"
BT_DB_NAME = "backtests.db"
TRADES_PER_PAGE = 5

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
            asset TEXT,
            direction TEXT,
            entry_price REAL,
            exit_price REAL,
            link_chart TEXT
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

def get_trades_filtered(user_id, result_filter=None, asset_filter=None, date_filter=None):
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
    query += " ORDER BY trade_date DESC"
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

def delete_trade(trade_id, user_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM trades WHERE id = ? AND user_id = ?", (trade_id, user_id))
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

def clear_trades(user_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM trades WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

def save_backtest(user_id, period_start, period_end, timeframe, asset, direction, entry_price, exit_price, link_chart):
    conn = sqlite3.connect(BT_DB_NAME)
    conn.execute("""
        INSERT INTO backtests (user_id, period_start, period_end, timeframe, asset, direction, entry_price, exit_price, link_chart)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, period_start, period_end, timeframe, asset, direction, entry_price, exit_price, link_chart))
    conn.commit()
    conn.close()

def get_backtests(user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    df = pd.read_sql_query("SELECT * FROM backtests WHERE user_id = ? ORDER BY period_start DESC", conn, params=(user_id,))
    conn.close()
    return df

def get_backtest_by_id(bt_id, user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    cur = conn.cursor()
    cur.execute("SELECT * FROM backtests WHERE id = ? AND user_id = ?", (bt_id, user_id))
    row = cur.fetchone()
    conn.close()
    if row:
        cols = ['id', 'user_id', 'period_start', 'period_end', 'timeframe', 'asset', 'direction', 'entry_price', 'exit_price', 'link_chart']
        return dict(zip(cols, row))
    return None

def clear_backtests(user_id):
    conn = sqlite3.connect(BT_DB_NAME)
    conn.execute("DELETE FROM backtests WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

# ========== ГРАФИК ==========
def generate_equity_chart(df, user_id):
    if df.empty:
        return None
    df = df.sort_values('trade_date')
    cum_pnl = df['pnl'].cumsum()
    
    plt.figure(figsize=(10, 5))
    plt.plot(cum_pnl.values, marker='o', color='#2b6cb0', linewidth=2, markersize=4)
    plt.axhline(0, color='red', linestyle='--', linewidth=1)
    plt.fill_between(range(len(cum_pnl)), 0, cum_pnl.values, where=(cum_pnl.values >= 0), color='green', alpha=0.3)
    plt.fill_between(range(len(cum_pnl)), 0, cum_pnl.values, where=(cum_pnl.values < 0), color='red', alpha=0.3)
    plt.title("📈 Кривая доходности", fontsize=14, fontweight='bold')
    plt.xlabel("Номер сделки")
    plt.ylabel("Накопленный P&L ($)")
    plt.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    
    fname = f"equity_{user_id}.png"
    plt.savefig(fname, bbox_inches='tight', dpi=100)
    plt.close()
    return fname

# ========== СТАТИСТИКА ==========
def get_stats_text(df):
    if df.empty:
        return "📭 Нет данных."
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
    emotion_text = "\n".join([f"{e}: {c}" for e, c in emotions.items()]) if emotions else "нет данных"
    return (
        f"📊 **Ваша статистика**\n\n"
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

def get_stats_text_short(df, title):
    if df.empty:
        return f"{title}\n\n📭 Нет данных."
    total = len(df)
    wins = len(df[df['pnl'] > 0])
    losses = len(df[df['pnl'] < 0])
    wr = wins/total*100 if total else 0
    total_pnl = df['pnl'].sum()
    return (
        f"{title}\n\n"
        f"📋 Сделок: {total}\n"
        f"✅ Тейков: {wins} | ❌ Стопов: {losses}\n"
        f"🎯 Винрейт: {wr:.1f}%\n"
        f"💰 P&L: ${total_pnl:.2f}"
    )

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
    df_exp = df_exp[['period_start', 'period_end', 'timeframe', 'asset', 'direction', 'entry_price', 'exit_price', 'link_chart']]
    df_exp.columns = ['📅 Начало', '📅 Конец', '⏱ Таймфрейм', '🪙 Актив', '📈 Направление', '💰 Вход', '💰 Выход', '🔗 Ссылка']
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

# ========== КЛАВИАТУРЫ ==========
def main_menu(lang):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Реальная торговля", callback_data="mode_real")],
        [InlineKeyboardButton(text="🔄 Бэктест", callback_data="mode_backtest")],
        [InlineKeyboardButton(text="⚙️ Настройки", callback_data="settings_menu")]
    ])

def real_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Сделка", callback_data="add_trade")],
        [InlineKeyboardButton(text="📋 Список сделок", callback_data="list_trades")],
        [InlineKeyboardButton(text="📊 Статистика", callback_data="real_stats_show")],
        [InlineKeyboardButton(text="📎 Excel", callback_data="real_excel")],
        [InlineKeyboardButton(text="🗑 Очистить всё", callback_data="real_clear")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def backtest_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Бэктест", callback_data="add_backtest")],
        [InlineKeyboardButton(text="📋 Список", callback_data="list_backtests")],
        [InlineKeyboardButton(text="📊 Статистика", callback_data="backtest_stats_show")],
        [InlineKeyboardButton(text="📎 Excel", callback_data="backtest_excel")],
        [InlineKeyboardButton(text="🗑 Очистить всё", callback_data="backtest_clear")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def settings_menu_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🌐 Сменить язык", callback_data="change_lang")],
        [InlineKeyboardButton(text="📞 Поддержка", callback_data="support")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def lang_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru")],
        [InlineKeyboardButton(text="🇬🇧 English", callback_data="lang_en")]
    ])

def direction_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🟢 LONG", callback_data="dir_LONG"),
         InlineKeyboardButton(text="🔴 SHORT", callback_data="dir_SHORT")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def result_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Тейк", callback_data="res_TAKE"),
         InlineKeyboardButton(text="❌ Стоп", callback_data="res_STOP"),
         InlineKeyboardButton(text="⚖️ БУ", callback_data="res_BU")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def emotion_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="😊 Спокойствие", callback_data="em_calm")],
        [InlineKeyboardButton(text="😨 Страх", callback_data="em_fear")],
        [InlineKeyboardButton(text="😈 Жадность", callback_data="em_greed")],
        [InlineKeyboardButton(text="🤬 Тильт", callback_data="em_tilt")],
        [InlineKeyboardButton(text="😌 Уверенность", callback_data="em_confidence")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def yesno_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Да", callback_data="yes")],
        [InlineKeyboardButton(text="❌ Нет", callback_data="no")]
    ])

def back_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def confirm_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚠️ ДА, УДАЛИТЬ", callback_data="clear_yes")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def trades_list_kb(trades, page, total_pages):
    buttons = []
    for _, row in trades.iterrows():
        pnl = row['pnl']
        emoji = "✅" if pnl > 0 else ("❌" if pnl < 0 else "⚖️")
        buttons.append([InlineKeyboardButton(text=f"{row['asset']} {emoji} ${pnl:.0f}", callback_data=f"view_{row['id']}")])
    
    nav = []
    if page > 1:
        nav.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"page_{page-1}"))
    if page < total_pages:
        nav.append(InlineKeyboardButton(text="Вперед ➡️", callback_data=f"page_{page+1}"))
    if nav:
        buttons.append(nav)
    
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="back")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def view_trade_kb(trade_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🗑 Удалить", callback_data=f"del_{trade_id}")],
        [InlineKeyboardButton(text="🔙 К списку", callback_data="list_trades")]
    ])

def list_backtests_kb(backtests):
    buttons = []
    for _, row in backtests.iterrows():
        buttons.append([InlineKeyboardButton(text=f"{row['asset']} | {row['period_start']}", callback_data=f"btview_{row['id']}")])
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="back")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def stats_main_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📈 Вся статистика", callback_data="stats_all")],
        [InlineKeyboardButton(text="💰 По активам", callback_data="stats_by_asset")],
        [InlineKeyboardButton(text="📅 По дате", callback_data="stats_by_date")],
        [InlineKeyboardButton(text="😊 По эмоциям", callback_data="stats_by_emotion")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back")]
    ])

def stats_assets_kb(assets):
    buttons = [[InlineKeyboardButton(text=a, callback_data=f"stats_asset_{a}")] for a in assets]
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="real_stats_show")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def stats_date_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📆 День", callback_data="stats_date_day")],
        [InlineKeyboardButton(text="📅 Неделя", callback_data="stats_date_week")],
        [InlineKeyboardButton(text="📊 Месяц", callback_data="stats_date_month")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="real_stats_show")]
    ])

def stats_emotions_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="😊 Спокойствие", callback_data="stats_em_calm")],
        [InlineKeyboardButton(text="😨 Страх", callback_data="stats_em_fear")],
        [InlineKeyboardButton(text="😈 Жадность", callback_data="stats_em_greed")],
        [InlineKeyboardButton(text="🤬 Тильт", callback_data="stats_em_tilt")],
        [InlineKeyboardButton(text="😌 Уверенность", callback_data="stats_em_confidence")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="real_stats_show")]
    ])

def filter_menu_kb(result_filter, asset_filter, date_filter, has_assets):
    buttons = []
    buttons.append([InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'all' else ''}Все", callback_data="filter_all")])
    buttons.append([InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'take' else ''}✅ Тейк", callback_data="filter_take")])
    buttons.append([InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'stop' else ''}❌ Стоп", callback_data="filter_stop")])
    buttons.append([InlineKeyboardButton(text=f"{'✅ ' if result_filter == 'bu' else ''}⚖️ БУ", callback_data="filter_bu")])
    if has_assets:
        buttons.append([InlineKeyboardButton(text="💰 По активу", callback_data="filter_asset_menu")])
    buttons.append([InlineKeyboardButton(text="📅 По дате", callback_data="filter_date_menu")])
    buttons.append([InlineKeyboardButton(text="🗑 Сбросить фильтры", callback_data="filter_clear")])
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="list_trades")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def filter_asset_kb(assets):
    buttons = [[InlineKeyboardButton(text=a, callback_data=f"filter_asset_{a}")] for a in assets]
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="filter_menu")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def filter_date_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📆 День", callback_data="filter_date_day")],
        [InlineKeyboardButton(text="📅 Неделя", callback_data="filter_date_week")],
        [InlineKeyboardButton(text="📊 Месяц", callback_data="filter_date_month")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="filter_menu")]
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
    asset = State()
    direction = State()
    entry_price = State()
    exit_price = State()
    link_chart = State()

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
    print(f"🌐 Веб-сервер запущен на порту {port}")
    await asyncio.Event().wait()

# ========== ОБРАБОТЧИКИ ==========
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# ---------- СТАРТ ----------
@dp.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await state.clear()
    uid = msg.from_user.id
    lang = get_user_lang(uid)
    if not lang:
        await msg.answer("🌐 Выберите язык / Choose language:", reply_markup=lang_kb())
        return
    await msg.answer("🎛 **Выберите режим работы:**", parse_mode="Markdown", reply_markup=main_menu(lang))

@dp.callback_query(F.data.startswith("lang_"))
async def set_lang(call: CallbackQuery, state: FSMContext):
    lang = call.data.split("_")[1]
    set_user_lang(call.from_user.id, lang)
    await call.message.delete()
    await call.message.answer("🎛 **Выберите режим работы:**", parse_mode="Markdown", reply_markup=main_menu(lang))
    await call.answer()

@dp.callback_query(F.data == "mode_real")
async def mode_real(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await call.message.edit_text("📊 **Реальная торговля**\n\nВыберите действие:", parse_mode="Markdown", reply_markup=real_menu())
    await call.answer()

@dp.callback_query(F.data == "mode_backtest")
async def mode_backtest(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await call.message.edit_text("🔄 **Бэктест**\n\nВыберите действие:", parse_mode="Markdown", reply_markup=backtest_menu())
    await call.answer()

@dp.callback_query(F.data == "back")
async def back(call: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = get_user_lang(call.from_user.id)
    await call.message.edit_text("🎛 **Выберите режим работы:**", parse_mode="Markdown", reply_markup=main_menu(lang))
    await call.answer()

# ---------- НАСТРОЙКИ ----------
@dp.callback_query(F.data == "settings_menu")
async def settings_menu(call: CallbackQuery):
    await call.message.edit_text("⚙️ **Настройки**\n\nВыберите действие:", parse_mode="Markdown", reply_markup=settings_menu_kb())
    await call.answer()

@dp.callback_query(F.data == "change_lang")
async def change_lang(call: CallbackQuery):
    await call.message.edit_text("🌐 **Выберите язык / Choose language:**", parse_mode="Markdown", reply_markup=lang_kb())
    await call.answer()

@dp.callback_query(F.data == "support")
async def support(call: CallbackQuery):
    await call.message.edit_text("📞 **Поддержка**\n\nПо вопросам пишите: @ваш_username", parse_mode="Markdown", reply_markup=settings_menu_kb())
    await call.answer()

# ---------- ДОБАВЛЕНИЕ СДЕЛКИ ----------
@dp.callback_query(F.data == "add_trade")
async def add_trade(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(TradeForm.asset)
    await call.message.edit_text("📝 Введите тикер (BTC, ETH, TON, AAPL):", reply_markup=back_kb())
    await call.answer()

@dp.message(TradeForm.asset)
async def get_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(TradeForm.direction)
    await msg.answer("📈 Выберите направление:", reply_markup=direction_kb())

@dp.callback_query(F.data == "dir_LONG" or F.data == "dir_SHORT")
async def get_direction(call: CallbackQuery, state: FSMContext):
    direction = "LONG" if call.data == "dir_LONG" else "SHORT"
    await state.update_data(direction=direction)
    await state.set_state(TradeForm.entry_price)
    await call.message.edit_text("💰 Введите цену входа:", reply_markup=back_kb())
    await call.answer()

@dp.message(TradeForm.entry_price)
async def get_entry(msg: Message, state: FSMContext):
    try:
        val = float(msg.text.replace(",", "."))
        await state.update_data(entry_price=val)
        await state.set_state(TradeForm.exit_price)
        await msg.answer("💰 Введите цену выхода:", reply_markup=back_kb())
    except:
        await msg.answer("❌ Ошибка! Введите число.", reply_markup=back_kb())

@dp.message(TradeForm.exit_price)
async def get_exit(msg: Message, state: FSMContext):
    try:
        val = float(msg.text.replace(",", "."))
        await state.update_data(exit_price=val)
        await state.set_state(TradeForm.volume)
        await msg.answer("📊 Введите объём позиции:", reply_markup=back_kb())
    except:
        await msg.answer("❌ Ошибка! Введите число.", reply_markup=back_kb())

@dp.message(TradeForm.volume)
async def get_volume(msg: Message, state: FSMContext):
    try:
        val = float(msg.text.replace(",", "."))
        await state.update_data(volume=val)
        await state.set_state(TradeForm.result)
        await msg.answer("🎯 Как закрылась сделка?", reply_markup=result_kb())
    except:
        await msg.answer("❌ Ошибка! Введите число.", reply_markup=back_kb())

@dp.callback_query(F.data.startswith("res_"))
async def get_result(call: CallbackQuery, state: FSMContext):
    result = call.data.split("_")[1]
    await state.update_data(result=result)
    await state.set_state(TradeForm.comment)
    await call.message.edit_text("📝 Введите комментарий (отправьте '-' чтобы пропустить):", reply_markup=back_kb())
    await call.answer()

@dp.message(TradeForm.comment)
async def get_comment(msg: Message, state: FSMContext):
    com = msg.text.strip()
    await state.update_data(comment="" if com == "-" else com)
    await state.set_state(TradeForm.add_link)
    await msg.answer("🔗 Хотите добавить ссылку на график?", reply_markup=yesno_kb())

@dp.callback_query(F.data == "yes")
async def add_link_yes(call: CallbackQuery, state: FSMContext):
    await state.set_state(TradeForm.link_url)
    await call.message.edit_text("🔗 Отправьте ссылку:", reply_markup=back_kb())
    await call.answer()

@dp.callback_query(F.data == "no")
async def add_link_no(call: CallbackQuery, state: FSMContext):
    await state.update_data(links="")
    await state.set_state(TradeForm.trade_date)
    await call.message.edit_text("📅 Введите дату (ДД.ММ.ГГГГ) или 'сегодня':", reply_markup=back_kb())
    await call.answer()

@dp.message(TradeForm.link_url)
async def get_link(msg: Message, state: FSMContext):
    await state.update_data(link_url=msg.text)
    await state.set_state(TradeForm.link_tf)
    await msg.answer("⏱ Какой это таймфрейм? (15м, 1ч, 4ч, 1д, 1н, 1м):", reply_markup=back_kb())

@dp.message(TradeForm.link_tf)
async def get_tf(msg: Message, state: FSMContext):
    tf = msg.text
    data = await state.get_data()
    links = data.get("links", "")
    new_link = f"{tf}: {data.get('link_url')}"
    links = f"{links}\n{new_link}" if links else new_link
    await state.update_data(links=links)
    await state.set_state(TradeForm.add_link)
    await msg.answer("✅ Ссылка сохранена! Добавить ещё?", reply_markup=yesno_kb())

@dp.message(TradeForm.trade_date)
async def get_date(msg: Message, state: FSMContext):
    dstr = msg.text.strip().lower()
    if dstr in ["сегодня", "today"]:
        trade_date = datetime.now().strftime("%Y-%m-%d")
    else:
        try:
            trade_date = datetime.strptime(dstr, "%d.%m.%Y").strftime("%Y-%m-%d")
        except:
            await msg.answer("❌ Ошибка! Введите дату в формате ДД.ММ.ГГГГ", reply_markup=back_kb())
            return
    await state.update_data(trade_date=trade_date)
    await state.set_state(TradeForm.emotion)
    await msg.answer("😊 Какие эмоции были?", reply_markup=emotion_kb())

@dp.callback_query(F.data.startswith("em_"))
async def get_emotion(call: CallbackQuery, state: FSMContext):
    emotion_map = {
        "calm": "😊 Спокойствие",
        "fear": "😨 Страх",
        "greed": "😈 Жадность",
        "tilt": "🤬 Тильт",
        "confidence": "😌 Уверенность"
    }
    emotion = emotion_map.get(call.data.split("_")[1], "😊 Спокойствие")
    data = await state.get_data()
    direction = data['direction']
    pnl = (data['exit_price'] - data['entry_price']) * data['volume'] if direction == "LONG" else (data['entry_price'] - data['exit_price']) * data['volume']
    if data['result'] == "BU":
        pnl = 0
    save_trade(
        user_id=call.from_user.id,
        asset=data['asset'], direction=direction,
        entry_price=data['entry_price'], exit_price=data['exit_price'],
        volume=data['volume'], pnl=pnl, result=data['result'],
        comment=data['comment'], trade_date=data['trade_date'],
        links=data.get('links', ''), emotion=emotion
    )
    await state.clear()
    await call.message.edit_text("✅ Сделка сохранена!", reply_markup=real_menu())
    await call.answer()

# ---------- СПИСОК СДЕЛОК С ФИЛЬТРАМИ И ПАГИНАЦИЕЙ ----------
@dp.callback_query(F.data == "list_trades")
async def list_trades(call: CallbackQuery, state: FSMContext):
    await state.update_data(page=1, result_filter="all", asset_filter=None, date_filter=None)
    await show_trades_page(call, state)

async def show_trades_page(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    page = data.get('page', 1)
    result_filter = data.get('result_filter', 'all')
    asset_filter = data.get('asset_filter', None)
    date_filter = data.get('date_filter', None)
    
    df = get_trades_filtered(call.from_user.id, result_filter, asset_filter, date_filter)
    if df.empty:
        await call.answer("📭 Нет данных", show_alert=True)
        return
    
    total = len(df)
    total_pages = (total + TRADES_PER_PAGE - 1) // TRADES_PER_PAGE
    if page > total_pages:
        page = total_pages
    start = (page - 1) * TRADES_PER_PAGE
    end = start + TRADES_PER_PAGE
    trades_df = df.iloc[start:end]
    
    await state.update_data(page=page)
    
    filter_text = ""
    if result_filter != "all":
        filter_text += f" | Фильтр: {result_filter}"
    if asset_filter:
        filter_text += f" | {asset_filter}"
    if date_filter:
        filter_text += f" | {date_filter}"
    
    text = f"📋 **Сделки** (страница {page}/{total_pages}){filter_text}"
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=trades_list_kb(trades_df, page, total_pages))
    await call.answer()

@dp.callback_query(F.data.startswith("page_"))
async def change_page(call: CallbackQuery, state: FSMContext):
    page = int(call.data.split("_")[1])
    await state.update_data(page=page)
    await show_trades_page(call, state)

@dp.callback_query(F.data == "filter_menu")
async def filter_menu(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    result_filter = data.get('result_filter', 'all')
    asset_filter = data.get('asset_filter', None)
    date_filter = data.get('date_filter', None)
    assets = get_all_assets(call.from_user.id)
    await call.message.edit_text("🔍 **Фильтры:**", parse_mode="Markdown", reply_markup=filter_menu_kb(result_filter, asset_filter, date_filter, len(assets) > 0))
    await call.answer()

@dp.callback_query(F.data.startswith("filter_"))
async def apply_filter(call: CallbackQuery, state: FSMContext):
    action = call.data.split("_")[1]
    
    if action == "all":
        await state.update_data(result_filter="all", asset_filter=None, date_filter=None, page=1)
    elif action == "take":
        await state.update_data(result_filter="take", page=1)
    elif action == "stop":
        await state.update_data(result_filter="stop", page=1)
    elif action == "bu":
        await state.update_data(result_filter="bu", page=1)
    elif action == "clear":
        await state.update_data(result_filter="all", asset_filter=None, date_filter=None, page=1)
        await show_trades_page(call, state)
        return
    elif action == "asset":
        await call.message.edit_text("💰 **Выберите актив:**", parse_mode="Markdown", reply_markup=filter_asset_kb(get_all_assets(call.from_user.id)))
        return
    elif action == "date":
        await call.message.edit_text("📅 **Выберите период:**", parse_mode="Markdown", reply_markup=filter_date_kb())
        return
    
    await show_trades_page(call, state)

@dp.callback_query(F.data.startswith("filter_asset_"))
async def apply_asset_filter(call: CallbackQuery, state: FSMContext):
    asset = call.data.split("_")[2]
    await state.update_data(asset_filter=asset, page=1)
    await show_trades_page(call, state)

@dp.callback_query(F.data.startswith("filter_date_"))
async def apply_date_filter(call: CallbackQuery, state: FSMContext):
    date_filter = call.data.split("_")[2]
    await state.update_data(date_filter=date_filter, page=1)
    await show_trades_page(call, state)

@dp.callback_query(F.data.startswith("view_"))
async def view_trade(call: CallbackQuery):
    trade_id = int(call.data.split("_")[1])
    trade = get_trade_by_id(trade_id, call.from_user.id)
    if not trade:
        await call.answer("Сделка не найдена", show_alert=True)
        return
    links = trade.get('links', '') or '-'
    dir_emoji = "🟢" if trade['direction'] == "LONG" else "🔴"
    result_text = {"TAKE": "✅ Тейк", "STOP": "❌ Стоп", "BU": "⚖️ БУ"}.get(trade['result'], trade['result'])
    text = (
        f"📋 **Сделка #{trade['id']}**\n\n"
        f"🪙 Актив: {trade['asset']}\n"
        f"📈 Направление: {dir_emoji} {trade['direction']}\n"
        f"💰 Вход: ${trade['entry_price']}\n"
        f"💰 Выход: ${trade['exit_price']}\n"
        f"📊 Объём: {trade['volume']}\n"
        f"💵 P&L: ${trade['pnl']}\n"
        f"🎯 Исход: {result_text}\n"
        f"📅 Дата: {trade['trade_date']}\n"
        f"😊 Эмоции: {trade['emotion']}\n"
        f"🔗 Ссылки:\n{links}\n"
        f"📝 Комментарий: {trade['comment'] or '-'}"
    )
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=view_trade_kb(trade_id))
    await call.answer()

# ---------- УДАЛЕНИЕ ----------
@dp.callback_query(F.data.startswith("del_"))
async def delete_confirm(call: CallbackQuery, state: FSMContext):
    trade_id = int(call.data.split("_")[1])
    await state.update_data(delete_id=trade_id)
    await call.message.edit_text(f"⚠️ Удалить сделку #{trade_id}?", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Да", callback_data="del_yes"),
         InlineKeyboardButton(text="❌ Нет", callback_data="list_trades")]
    ]))
    await call.answer()

@dp.callback_query(F.data == "del_yes")
async def delete_execute(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    trade_id = data.get('delete_id')
    if trade_id:
        delete_trade(trade_id, call.from_user.id)
    await state.clear()
    await call.message.edit_text("🗑 Сделка удалена!", reply_markup=real_menu())
    await call.answer()

# ---------- СТАТИСТИКА (меню) ----------
@dp.callback_query(F.data == "real_stats_show")
async def real_stats_menu(call: CallbackQuery):
    await call.message.edit_text("📊 **Выберите тип статистики:**", parse_mode="Markdown", reply_markup=stats_main_kb())
    await call.answer()

@dp.callback_query(F.data == "stats_all")
async def stats_all(call: CallbackQuery):
    df = get_trades_filtered(call.from_user.id)
    text = get_stats_text(df)
    chart = generate_equity_chart(df, call.from_user.id)
    if chart:
        await call.message.answer_photo(photo=FSInputFile(chart), caption=text, parse_mode="Markdown")
        os.remove(chart)
    else:
        await call.message.answer(text, parse_mode="Markdown")
    await call.message.answer("📊 **Выберите тип статистики:**", parse_mode="Markdown", reply_markup=stats_main_kb())
    await call.answer()

@dp.callback_query(F.data == "stats_by_asset")
async def stats_by_asset_menu(call: CallbackQuery):
    assets = get_all_assets(call.from_user.id)
    if not assets:
        await call.answer("📭 Нет активов", show_alert=True)
        return
    await call.message.edit_text("💰 **Выберите актив:**", parse_mode="Markdown", reply_markup=stats_assets_kb(assets))
    await call.answer()

@dp.callback_query(F.data.startswith("stats_asset_"))
async def stats_asset_show(call: CallbackQuery):
    asset = call.data.split("_")[2]
    df = get_trades_filtered(call.from_user.id, asset_filter=asset)
    text = get_stats_text(df)
    chart = generate_equity_chart(df, call.from_user.id)
    if chart:
        await call.message.answer_photo(photo=FSInputFile(chart), caption=text, parse_mode="Markdown")
        os.remove(chart)
    else:
        await call.message.answer(text, parse_mode="Markdown")
    await call.message.answer("💰 **Выберите актив:**", parse_mode="Markdown", reply_markup=stats_assets_kb(get_all_assets(call.from_user.id)))
    await call.answer()

@dp.callback_query(F.data == "stats_by_date")
async def stats_by_date_menu(call: CallbackQuery):
    await call.message.edit_text("📅 **Выберите период:**", parse_mode="Markdown", reply_markup=stats_date_kb())
    await call.answer()

@dp.callback_query(F.data.startswith("stats_date_"))
async def stats_date_show(call: CallbackQuery):
    period = call.data.split("_")[2]
    days = {"day": 1, "week": 7, "month": 30}
    start = (datetime.now() - timedelta(days=days[period])).strftime("%Y-%m-%d")
    df = get_trades_filtered(call.from_user.id, date_filter=period)
    titles = {"day": "За сегодня", "week": "За неделю", "month": "За месяц"}
    text = get_stats_text_short(df, titles.get(period, "Статистика"))
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_date_kb())
    await call.answer()

@dp.callback_query(F.data == "stats_by_emotion")
async def stats_by_emotion_menu(call: CallbackQuery):
    await call.message.edit_text("😊 **Выберите эмоцию:**", parse_mode="Markdown", reply_markup=stats_emotions_kb())
    await call.answer()

@dp.callback_query(F.data.startswith("stats_em_"))
async def stats_emotion_show(call: CallbackQuery):
    em_map = {
        "calm": "😊 Спокойствие",
        "fear": "😨 Страх",
        "greed": "😈 Жадность",
        "tilt": "🤬 Тильт",
        "confidence": "😌 Уверенность"
    }
    emotion = em_map.get(call.data.split("_")[2], "😊 Спокойствие")
    df = get_trades_filtered(call.from_user.id)
    df = df[df['emotion'] == emotion]
    text = get_stats_text(df)
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=stats_emotions_kb())
    await call.answer()

# ---------- БЭКТЕСТ СТАТИСТИКА ----------
@dp.callback_query(F.data == "backtest_stats_show")
async def backtest_stats_show(call: CallbackQuery):
    df = get_backtests(call.from_user.id)
    if df.empty:
        await call.answer("📭 Нет данных", show_alert=True)
        return
    total = len(df)
    text = f"📊 **Статистика бэктестов**\n\n📋 Всего бэктестов: {total}"
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=backtest_menu())
    await call.answer()

# ---------- EXCEL ----------
@dp.callback_query(F.data == "real_excel")
async def real_excel(call: CallbackQuery):
    df = get_trades_filtered(call.from_user.id)
    if df.empty:
        await call.answer("📭 Нет данных", show_alert=True)
        return
    fname = export_real_to_excel(df, call.from_user.id)
    await call.message.answer_document(document=FSInputFile(fname), caption="📊 Ваш отчёт (реальная торговля)")
    os.remove(fname)
    await call.answer()

@dp.callback_query(F.data == "backtest_excel")
async def backtest_excel(call: CallbackQuery):
    df = get_backtests(call.from_user.id)
    if df.empty:
        await call.answer("📭 Нет данных", show_alert=True)
        return
    fname = export_backtest_to_excel(df, call.from_user.id)
    await call.message.answer_document(document=FSInputFile(fname), caption="📊 Ваш отчёт (бэктест)")
    os.remove(fname)
    await call.answer()

# ---------- ОЧИСТКА ----------
@dp.callback_query(F.data == "real_clear")
async def real_clear_confirm(call: CallbackQuery):
    await call.message.edit_text("⚠️ Удалить ВСЕ сделки реальной торговли?", reply_markup=confirm_kb())
    await call.answer()

@dp.callback_query(F.data == "backtest_clear")
async def backtest_clear_confirm(call: CallbackQuery):
    await call.message.edit_text("⚠️ Удалить ВСЕ бэктесты?", reply_markup=confirm_kb())
    await call.answer()

@dp.callback_query(F.data == "clear_yes")
async def clear_yes(call: CallbackQuery):
    clear_trades(call.from_user.id)
    await call.message.edit_text("🗑 Журнал реальной торговли очищен!", reply_markup=real_menu())
    await call.answer()

# ---------- БЭКТЕСТ ----------
@dp.callback_query(F.data == "add_backtest")
async def add_backtest(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(BacktestForm.period_start)
    await call.message.edit_text("📅 Введите НАЧАЛО периода (ДД.ММ.ГГГГ):", reply_markup=back_kb())
    await call.answer()

@dp.message(BacktestForm.period_start)
async def bt_start(msg: Message, state: FSMContext):
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_start=d)
        await state.set_state(BacktestForm.period_end)
        await msg.answer("📅 Введите КОНЕЦ периода (ДД.ММ.ГГГГ):", reply_markup=back_kb())
    except:
        await msg.answer("❌ Ошибка! Введите дату в формате ДД.ММ.ГГГГ", reply_markup=back_kb())

@dp.message(BacktestForm.period_end)
async def bt_end(msg: Message, state: FSMContext):
    try:
        d = datetime.strptime(msg.text.strip(), "%d.%m.%Y").strftime("%Y-%m-%d")
        await state.update_data(period_end=d)
        await state.set_state(BacktestForm.timeframe)
        await msg.answer("⏱ Введите таймфрейм (M5, H1, H4, D1, W1):", reply_markup=back_kb())
    except:
        await msg.answer("❌ Ошибка! Введите дату в формате ДД.ММ.ГГГГ", reply_markup=back_kb())

@dp.message(BacktestForm.timeframe)
async def bt_tf(msg: Message, state: FSMContext):
    await state.update_data(timeframe=msg.text.upper())
    await state.set_state(BacktestForm.asset)
    await msg.answer("📝 Введите тикер (BTC, ETH, TON, AAPL):", reply_markup=back_kb())

@dp.message(BacktestForm.asset)
async def bt_asset(msg: Message, state: FSMContext):
    await state.update_data(asset=msg.text.upper())
    await state.set_state(BacktestForm.direction)
    await msg.answer("📈 Выберите направление:", reply_markup=direction_kb())

@dp.callback_query(F.data == "dir_LONG" or F.data == "dir_SHORT")
async def bt_direction(call: CallbackQuery, state: FSMContext):
    direction = "LONG" if call.data == "dir_LONG" else "SHORT"
    await state.update_data(direction=direction)
    await state.set_state(BacktestForm.entry_price)
    await call.message.edit_text("💰 Введите цену входа:", reply_markup=back_kb())
    await call.answer()

@dp.message(BacktestForm.entry_price)
async def bt_entry(msg: Message, state: FSMContext):
    try:
        await state.update_data(entry_price=float(msg.text.replace(",", ".")))
        await state.set_state(BacktestForm.exit_price)
        await msg.answer("💰 Введите цену выхода:", reply_markup=back_kb())
    except:
        await msg.answer("❌ Ошибка! Введите число.", reply_markup=back_kb())

@dp.message(BacktestForm.exit_price)
async def bt_exit(msg: Message, state: FSMContext):
    try:
        exit_p = float(msg.text.replace(",", "."))
        await state.update_data(exit_price=exit_p)
        await state.set_state(BacktestForm.link_chart)
        await msg.answer("🔗 Ссылка на скриншот (0 если нет):", reply_markup=back_kb())
    except:
        await msg.answer("❌ Ошибка! Введите число.", reply_markup=back_kb())

@dp.message(BacktestForm.link_chart)
async def bt_link(msg: Message, state: FSMContext):
    link = msg.text if msg.text != "0" else "-"
    data = await state.get_data()
    save_backtest(
        user_id=msg.from_user.id,
        period_start=data['period_start'],
        period_end=data['period_end'],
        timeframe=data['timeframe'],
        asset=data['asset'],
        direction=data['direction'],
        entry_price=data['entry_price'],
        exit_price=data['exit_price'],
        link_chart=link
    )
    await state.clear()
    await msg.answer("✅ Бэктест сохранён!", reply_markup=backtest_menu())

@dp.callback_query(F.data == "list_backtests")
async def list_backtests(call: CallbackQuery):
    df = get_backtests(call.from_user.id)
    if df.empty:
        await call.answer("📭 Нет данных", show_alert=True)
        return
    await call.message.edit_text("📋 **Список бэктестов:**", parse_mode="Markdown", reply_markup=list_backtests_kb(df))
    await call.answer()

@dp.callback_query(F.data.startswith("btview_"))
async def view_backtest(call: CallbackQuery):
    bt_id = int(call.data.split("_")[1])
    bt = get_backtest_by_id(bt_id, call.from_user.id)
    if not bt:
        await call.answer("Бэктест не найден", show_alert=True)
        return
    text = (
        f"📋 **Бэктест #{bt['id']}**\n\n"
        f"🪙 Актив: {bt['asset']}\n"
        f"📈 Направление: {'🟢 LONG' if bt['direction'] == 'LONG' else '🔴 SHORT'}\n"
        f"💰 Вход: ${bt['entry_price']}\n"
        f"💰 Выход: ${bt['exit_price']}\n"
        f"📅 Период: {bt['period_start']} — {bt['period_end']}\n"
        f"⏱ Таймфрейм: {bt['timeframe']}\n"
        f"🔗 Ссылка: {bt['link_chart']}"
    )
    await call.message.edit_text(text, parse_mode="Markdown", reply_markup=backtest_menu())
    await call.answer()

# ---------- КОМАНДЫ ----------
@dp.message(Command("new"))
async def cmd_new(msg: Message, state: FSMContext):
    await state.clear()
    await state.set_state(TradeForm.asset)
    await msg.answer("📝 Введите тикер (BTC, ETH, TON, AAPL):", reply_markup=back_kb())

@dp.message(Command("stats"))
async def cmd_stats(msg: Message):
    df = get_trades_filtered(msg.from_user.id)
    text = get_stats_text(df)
    chart = generate_equity_chart(df, msg.from_user.id)
    if chart:
        await msg.answer_photo(photo=FSInputFile(chart), caption=text, parse_mode="Markdown")
        os.remove(chart)
    else:
        await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("day"))
async def cmd_day(msg: Message):
    today = datetime.now().strftime("%Y-%m-%d")
    df = get_trades_filtered(msg.from_user.id, date_filter="day")
    text = get_stats_text_short(df, "📆 Статистика за сегодня")
    await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("week"))
async def cmd_week(msg: Message):
    df = get_trades_filtered(msg.from_user.id, date_filter="week")
    text = get_stats_text_short(df, "📅 Статистика за неделю")
    await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("month"))
async def cmd_month(msg: Message):
    df = get_trades_filtered(msg.from_user.id, date_filter="month")
    text = get_stats_text_short(df, "📊 Статистика за месяц")
    await msg.answer(text, parse_mode="Markdown")

@dp.message(Command("clear"))
async def cmd_clear(msg: Message):
    clear_trades(msg.from_user.id)
    await msg.answer("🗑 Журнал очищен!")

@dp.message(Command("get_real"))
async def cmd_get_real(msg: Message):
    df = get_trades_filtered(msg.from_user.id)
    if df.empty:
        await msg.answer("📭 Нет данных")
        return
    fname = export_real_to_excel(df, msg.from_user.id)
    await msg.answer_document(document=FSInputFile(fname), caption="📊 Ваш отчёт (реальная торговля)")
    os.remove(fname)

@dp.message(Command("get_backtest"))
async def cmd_get_backtest(msg: Message):
    df = get_backtests(msg.from_user.id)
    if df.empty:
        await msg.answer("📭 Нет данных")
        return
    fname = export_backtest_to_excel(df, msg.from_user.id)
    await msg.answer_document(document=FSInputFile(fname), caption="📊 Ваш отчёт (бэктест)")
    os.remove(fname)

# ========== ЗАПУСК ==========
async def set_commands():
    await bot.set_my_commands([
        BotCommand(command="start", description="Главное меню"),
        BotCommand(command="new", description="➕ Новая сделка"),
        BotCommand(command="stats", description="📊 Вся статистика"),
        BotCommand(command="day", description="📆 За сегодня"),
        BotCommand(command="week", description="📅 За неделю"),
        BotCommand(command="month", description="📊 За месяц"),
        BotCommand(command="clear", description="🗑 Очистить журнал"),
        BotCommand(command="get_real", description="📎 Excel (реальная)"),
        BotCommand(command="get_backtest", description="📎 Excel (бэктест)"),
    ])

async def main():
    init_dbs()
    await set_commands()
    print("✅ Бот успешно запущен!")
    await asyncio.gather(
        run_web_server(),
        dp.start_polling(bot)
    )

if __name__ == "__main__":
    asyncio.run(main())
