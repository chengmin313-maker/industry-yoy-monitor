# -*- coding: utf-8 -*-
"""
================================================================
細產業營收 YoY 動能追蹤系統 v2.1
================================================================
功能：
  1. 讀取細產業分類 Excel（成分股細產業整理_YYYYMMDD.xlsx）
  2. 從 MOPS 公開資訊觀測站爬取月營收（上市+上櫃）
  3. 按子產業加總，計算每月 YoY
  4. ⭐ Top 5 動能排行榜（景氣循環股優先）
     - 近 3 月 YoY 平均排序
     - 入榜條件：最新月 YoY > 前月 YoY（仍在加速）
     - 淘汰條件：連續 2 個月 YoY 下滑 → 踢出，下一名遞補
  5. 🔥 YoY 負翻正偵測（前月 < 0 → 本月 > 0）
  6. 💯 YoY 破百偵測（本月 >= 100%）
  7. TG 推送（Top 5 + 負翻正 + 破百）

使用方式：
  python industry_yoy_monitor.py

需求：
  pip install requests openpyxl pandas beautifulsoup4 lxml
  同目錄下放「成分股細產業整理_YYYYMMDD.xlsx」

每月 11 號後執行（等營收公告完畢），或搭配排程自動化
================================================================
"""

import os
import re
import time
import json
import requests
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from collections import defaultdict

# ==========================================
# 系統設定
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_excel_candidates = sorted(
    [f for f in os.listdir(SCRIPT_DIR)
     if f.startswith("成分股細產業整理") and f.endswith(".xlsx")],
    reverse=True
)
EXCEL_FILE = os.path.join(SCRIPT_DIR, _excel_candidates[0]) if _excel_candidates else os.path.join(SCRIPT_DIR, "成分股細產業整理.xlsx")
CACHE_DIR  = os.path.join(SCRIPT_DIR, "revenue_cache")
CYCLICAL_JSON = os.path.join(SCRIPT_DIR, "cyclical_industries.json")

# 偵測參數
MIN_STOCKS = 5          # 子產業至少要有 N 檔股票才納入分析
DECEL_MONTHS = 2        # 連續 N 個月下滑 → 淘汰出 Top 5
FETCH_MONTHS = 18       # 往回爬 N 個月的營收資料
TOP_N = 5               # 動能排行榜名次

# Telegram 設定
TG_BOT_TOKEN = ""       # 填入你的 Bot Token
TG_CHAT_ID   = ""       # 填入你的 Chat ID

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


# ==========================================
# 0. 景氣循環股分類（JSON 管理）
# ==========================================
DEFAULT_CYCLICAL = {
    "_說明": "景氣循環股產業清單 — 程式自動比對細產業名稱（部分匹配）。可手動新增/刪除。",
    "_更新日期": "",
    "產業關鍵字": [
        "塑膠", "石化", "化工", "化學", "PVC", "PET", "ABS",
        "鋼鐵", "鋼", "鑄造", "銅", "鋁", "鎳", "鋅", "金屬",
        "紙漿", "紙", "造紙",
        "水泥", "玻璃", "陶瓷", "耐火材",
        "航運", "貨櫃", "散裝", "航空", "物流",
        "DRAM", "記憶體", "面板", "LCD", "OLED", "顯示器",
        "矽晶圓", "晶圓材料", "被動元件", "MLCC",
        "PCB", "銅箔基板", "CCL", "玻纖布",
        "營建", "建設", "建材", "不動產", "資產股",
        "石油", "天然氣", "煤", "能源", "太陽能",
        "紡織", "纖維", "橡膠", "輪胎",
        "汽車零組件", "工具機", "機械", "光通訊",
    ]
}


def load_cyclical_industries():
    if not os.path.exists(CYCLICAL_JSON):
        DEFAULT_CYCLICAL["_更新日期"] = datetime.now().strftime("%Y-%m-%d")
        with open(CYCLICAL_JSON, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CYCLICAL, f, ensure_ascii=False, indent=2)
        print(f"  📝 已建立景氣循環清單: {os.path.basename(CYCLICAL_JSON)}")
    with open(CYCLICAL_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)
    keywords = data.get("產業關鍵字", [])
    print(f"  🔄 景氣循環關鍵字: {len(keywords)} 個")
    return keywords


def is_cyclical(industry_name, cyclical_keywords):
    for kw in cyclical_keywords:
        if kw in industry_name:
            return True
    return False


# ==========================================
# 1. 讀取細產業分類表
# ==========================================
def load_industry_mapping():
    print(f"📋 讀取細產業分類表: {os.path.basename(EXCEL_FILE)}")
    wb = load_workbook(EXCEL_FILE, read_only=True)
    ws = wb['成分股細產業整理']

    stock_industries = {}
    industry_stocks = defaultdict(list)

    for row in ws.iter_rows(min_row=2, values_only=True):
        code     = str(row[0]).strip() if row[0] else ""
        code_raw = str(row[1]).strip() if row[1] else ""
        name     = str(row[2]).strip() if row[2] else ""
        ind_str  = str(row[3]).strip() if row[3] else ""

        if not code or not code.isdigit():
            continue

        market = "上櫃" if ".TWO" in code_raw else "上市"
        industries = [x.strip() for x in ind_str.split("/") if x.strip()]
        stock_industries[code] = {
            "name": name, "market": market, "industries": industries
        }
        for ind in industries:
            if code not in industry_stocks[ind]:
                industry_stocks[ind].append(code)

    wb.close()
    filtered = {k: v for k, v in industry_stocks.items() if len(v) >= MIN_STOCKS}
    print(f"  ✅ 共 {len(stock_industries)} 檔股票，{len(filtered)} 個子產業 (>= {MIN_STOCKS} 檔)")
    return stock_industries, filtered


# ==========================================
# 2. 爬取月營收（MOPS 公開資訊觀測站）
# ==========================================
MOPS_BASE = "https://mopsov.twse.com.tw"
MOPS_AJAX_BASE = "https://mopsov.twse.com.tw/mops/web"


def _parse_revenue_tables(dfs):
    result = {}
    for df in dfs:
        if df.shape[0] < 2 or df.shape[1] < 5:
            continue
        df.columns = range(len(df.columns))
        for _, row in df.iterrows():
            code = str(row[0]).strip()
            if not code or len(code) < 4 or len(code) > 6:
                continue
            if not code[0].isdigit():
                continue
            if code in ("合計", "公司代號", "公司", "合計:"):
                continue
            try:
                rev_raw = str(row[2]).replace(",", "").replace(" ", "").strip()
                rev = int(float(rev_raw))
                if rev > 0:
                    result[code] = rev
            except (ValueError, TypeError):
                continue
    return result


def _parse_revenue_bs4(html_text):
    result = {}
    soup = BeautifulSoup(html_text, "html.parser")
    for table in soup.find_all("table"):
        for row in table.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) < 3:
                continue
            code = cells[0].get_text(strip=True)
            if not code or not code[0].isdigit() or len(code) < 4 or len(code) > 6:
                continue
            rev_text = cells[2].get_text(strip=True).replace(",", "")
            try:
                rev = int(float(rev_text))
                if rev > 0:
                    result[code] = rev
            except (ValueError, TypeError):
                continue
    return result


def _fetch_and_parse(url, method="GET", payload=None, encoding="big5"):
    result = {}
    try:
        if method == "GET":
            res = requests.get(url, headers=HEADERS, timeout=20, verify=False)
        else:
            res = requests.post(url, data=payload, headers=HEADERS, timeout=25, verify=False)
        res.encoding = encoding
        if res.status_code == 200 and len(res.text) > 500:
            try:
                dfs = pd.read_html(StringIO(res.text))
                result = _parse_revenue_tables(dfs)
            except ValueError:
                result = _parse_revenue_bs4(res.text)
    except Exception:
        pass
    return result


def fetch_monthly_revenue(year, month, market="sii"):
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"rev_{market}_{year}_{month:02d}.json")
    if os.path.exists(cache_file):
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)

    roc_year = year - 1911
    result = {}

    for suffix in ["0", "1"]:
        url = f"{MOPS_BASE}/nas/t21/{market}/t21sc03_{roc_year}_{month}_{suffix}.html"
        parsed = _fetch_and_parse(url, "GET", encoding="big5")
        result.update(parsed)

    if not result:
        for suffix in ["0", "1"]:
            url = f"https://mops.twse.com.tw/nas/t21/{market}/t21sc03_{roc_year}_{month}_{suffix}.html"
            parsed = _fetch_and_parse(url, "GET", encoding="big5")
            result.update(parsed)

    if not result:
        for ep in ["ajax_t21sc03_ifrs", "ajax_t21sc03"]:
            if result:
                break
            payload = {
                "encodeURIComponent": "1", "step": "1", "firstin": "1",
                "off": "1", "TYPEK": market,
                "year": str(roc_year), "month": str(month),
            }
            result = _fetch_and_parse(f"{MOPS_AJAX_BASE}/{ep}", "POST", payload, "utf-8")

    if not result:
        return result
    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(result, f)
    return result


def fetch_all_revenue(year, month):
    sii = fetch_monthly_revenue(year, month, "sii")
    time.sleep(2)
    otc = fetch_monthly_revenue(year, month, "otc")
    time.sleep(1)
    return {**sii, **otc}


# ==========================================
# 3. 計算子產業合計營收與 YoY
# ==========================================
def compute_industry_yoy(industry_stocks, months_data):
    rows = []
    for ind, stocks in industry_stocks.items():
        for ym, rev_dict in months_data.items():
            total = sum(rev_dict.get(s, 0) for s in stocks)
            if total == 0:
                continue
            y, m = int(ym[:4]), int(ym[5:])
            last_year_ym = f"{y - 1}-{m:02d}"
            last_year_total = 0
            if last_year_ym in months_data:
                last_year_total = sum(months_data[last_year_ym].get(s, 0) for s in stocks)
            yoy = None
            if last_year_total > 0:
                yoy = round((total / last_year_total - 1) * 100, 2)
            rows.append({
                "子產業": ind, "年月": ym,
                "成分股數": len([s for s in stocks if rev_dict.get(s, 0) > 0]),
                "合計營收(千)": total, "去年同月(千)": last_year_total,
                "YoY%": yoy
            })
    return pd.DataFrame(rows)


# ==========================================
# 4. Top 5 動能排行榜
# ==========================================
def compute_top5(yoy_df, cyclical_keywords):
    """
    Top 5 動能排行榜
    排序: 近 3 月 YoY 平均（景氣循環股優先）
    入榜條件: 最新月 YoY > 前月 YoY（仍在加速）
    淘汰條件: 連續 2 個月 YoY 下滑 → 跳過
    """
    sorted_months = sorted(yoy_df["年月"].unique())
    if len(sorted_months) < 3:
        return []

    latest_3 = sorted_months[-3:]
    latest_1 = sorted_months[-1]
    prev_1   = sorted_months[-2]

    candidates = []

    for ind in yoy_df["子產業"].unique():
        ind_data = yoy_df[yoy_df["子產業"] == ind].sort_values("年月")
        ind_data = ind_data[ind_data["YoY%"].notna()]

        recent3 = ind_data[ind_data["年月"].isin(latest_3)]
        if len(recent3) < 2:
            continue

        avg_yoy_3m = recent3["YoY%"].mean()
        latest_yoy = ind_data[ind_data["年月"] == latest_1]["YoY%"].values
        prev_yoy   = ind_data[ind_data["年月"] == prev_1]["YoY%"].values

        if len(latest_yoy) == 0:
            continue
        latest_yoy = latest_yoy[0]
        prev_yoy   = prev_yoy[0] if len(prev_yoy) > 0 else None

        is_accelerating = (prev_yoy is not None) and (latest_yoy > prev_yoy)

        # 連續下滑月數
        yoy_vals = ind_data["YoY%"].tolist()
        consecutive_decline = 0
        if len(yoy_vals) >= 3:
            for i in range(len(yoy_vals) - 1, 0, -1):
                if yoy_vals[i] < yoy_vals[i - 1]:
                    consecutive_decline += 1
                else:
                    break

        is_cyc = is_cyclical(ind, cyclical_keywords)
        trend_vals = ind_data["YoY%"].tolist()[-4:]
        trend_str = " → ".join([f"{v:+.1f}" for v in trend_vals])
        stock_count = int(ind_data.iloc[-1]["成分股數"]) if len(ind_data) > 0 else 0

        candidates.append({
            "子產業": ind,
            "景氣循環": is_cyc,
            "近3月YoY均": round(avg_yoy_3m, 2),
            "最新YoY%": latest_yoy,
            "前月YoY%": prev_yoy,
            "加速中": is_accelerating,
            "連續下滑月數": consecutive_decline,
            "成分股數": stock_count,
            "近期趨勢": trend_str,
        })

    # 排序: 景氣循環優先 → 近 3 月 YoY 均高→低
    candidates.sort(key=lambda x: (not x["景氣循環"], -x["近3月YoY均"]))

    top5_list = []
    rank = 0
    for c in candidates:
        if c["連續下滑月數"] >= DECEL_MONTHS:
            continue
        if not c["加速中"]:
            continue
        rank += 1
        c["排名"] = rank
        top5_list.append(c)
        if rank >= TOP_N:
            break

    return top5_list


# ==========================================
# 5. YoY 負翻正 + YoY 破百 偵測
# ==========================================
def detect_signals(yoy_df, cyclical_keywords):
    """
    偵測兩種訊號：
      負翻正: 前月 YoY < 0 → 本月 YoY > 0
      破百:   本月 YoY >= 100%
    回傳 (neg_to_pos_list, over100_list)
    """
    sorted_months = sorted(yoy_df["年月"].unique())
    if len(sorted_months) < 2:
        return [], []

    latest_1 = sorted_months[-1]
    prev_1   = sorted_months[-2]

    neg_to_pos = []
    over_100   = []

    for ind in yoy_df["子產業"].unique():
        ind_data = yoy_df[yoy_df["子產業"] == ind].sort_values("年月")
        ind_data = ind_data[ind_data["YoY%"].notna()]

        latest_row = ind_data[ind_data["年月"] == latest_1]
        prev_row   = ind_data[ind_data["年月"] == prev_1]

        if len(latest_row) == 0:
            continue

        latest_yoy = latest_row["YoY%"].values[0]
        prev_yoy   = prev_row["YoY%"].values[0] if len(prev_row) > 0 else None
        stock_count = int(latest_row["成分股數"].values[0])
        is_cyc = is_cyclical(ind, cyclical_keywords)

        trend_vals = ind_data["YoY%"].tolist()[-4:]
        trend_str = " → ".join([f"{v:+.1f}" for v in trend_vals])

        info = {
            "子產業": ind,
            "景氣循環": is_cyc,
            "最新YoY%": latest_yoy,
            "前月YoY%": prev_yoy,
            "成分股數": stock_count,
            "近期趨勢": trend_str,
        }

        # 負翻正: 前月 < 0, 本月 > 0
        if prev_yoy is not None and prev_yoy < 0 and latest_yoy > 0:
            neg_to_pos.append(info)

        # 破百: 本月 >= 100%
        if latest_yoy >= 100:
            over_100.append(info)

    # 排序: 景氣循環優先 → YoY 高→低
    neg_to_pos.sort(key=lambda x: (not x["景氣循環"], -x["最新YoY%"]))
    over_100.sort(key=lambda x: (not x["景氣循環"], -x["最新YoY%"]))

    return neg_to_pos, over_100


# ==========================================
# 6. TG 推送
# ==========================================
RANK_EMOJI = {1: "🥇", 2: "🥈", 3: "🥉", 4: "4️⃣", 5: "5️⃣"}


def send_tg_message(msg):
    if not TG_BOT_TOKEN or not TG_CHAT_ID:
        return False
    url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg, "parse_mode": "HTML"}
    try:
        requests.post(url, data=payload, timeout=10)
        return True
    except Exception as e:
        print(f"  ❌ TG 推送失敗: {e}")
        return False


def send_tg_report(top5, neg_to_pos, over_100, latest_month):
    if not TG_BOT_TOKEN or not TG_CHAT_ID:
        print("⚠️ 未設定 TG Bot Token / Chat ID，跳過推送")
        return

    lines = []

    # ── Top 5 動能排行 ──
    lines.append(f"📊 <b>細產業 YoY 動能 Top {TOP_N}</b>  ({latest_month})")
    lines.append("")

    if top5:
        for t in top5:
            emoji = RANK_EMOJI.get(t["排名"], "▪️")
            cyc_tag = " 🔄" if t["景氣循環"] else ""
            lines.append(
                f"{emoji} <b>{t['子產業']}</b>{cyc_tag}  ({t['成分股數']}檔)\n"
                f"    YoY {t['最新YoY%']:+.1f}%  3M均 {t['近3月YoY均']:+.1f}%\n"
                f"    {t['近期趨勢']}"
            )
    else:
        lines.append("  （無符合條件的產業）")

    # ── YoY 負翻正 ──
    if neg_to_pos:
        lines.append("")
        lines.append(f"🔥 <b>YoY 負翻正</b> ({len(neg_to_pos)}個)")
        for n in neg_to_pos[:8]:
            cyc_tag = " 🔄" if n["景氣循環"] else ""
            lines.append(
                f"  ▸ {n['子產業']}{cyc_tag}  "
                f"{n['前月YoY%']:+.1f}% → <b>{n['最新YoY%']:+.1f}%</b>"
            )

    # ── YoY 破百 ──
    if over_100:
        lines.append("")
        lines.append(f"💯 <b>YoY ≥ 100%</b> ({len(over_100)}個)")
        for o in over_100[:8]:
            cyc_tag = " 🔄" if o["景氣循環"] else ""
            lines.append(
                f"  ▸ {o['子產業']}{cyc_tag}  "
                f"YoY <b>{o['最新YoY%']:+.1f}%</b>  ({o['成分股數']}檔)\n"
                f"    {o['近期趨勢']}"
            )

    lines.append("")
    lines.append("🔄=景氣循環  ▲=加速中")

    msg = "\n".join(lines)

    # TG 訊息上限 4096 字元
    if len(msg) > 4000:
        msg = msg[:3990] + "\n...(已截斷)"

    if send_tg_message(msg):
        print("📲 TG 推送完成")


# ==========================================
# 主流程
# ==========================================
def main():
    print("=" * 60)
    print("細產業營收 YoY 動能追蹤系統 v2.1")
    print(f"執行時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"MOPS 域名: {MOPS_BASE}")
    print("=" * 60)

    # 0. 載入景氣循環清單
    cyclical_keywords = load_cyclical_industries()

    # 1. 讀取產業分類
    stock_info, industry_stocks = load_industry_mapping()
    cyc_count = sum(1 for ind in industry_stocks if is_cyclical(ind, cyclical_keywords))
    print(f"  🔄 其中 {cyc_count} 個為景氣循環產業")

    # 2. 爬取月營收
    print(f"\n📡 開始爬取月營收 (往回 {FETCH_MONTHS} 個月)...")
    months_data = {}
    now = datetime.now()

    for i in range(FETCH_MONTHS):
        target = now - timedelta(days=30 * (i + 1))
        y, m = target.year, target.month
        ym_key = f"{y}-{m:02d}"
        print(f"  ⏳ {ym_key}...", end="", flush=True)
        rev = fetch_all_revenue(y, m)
        if rev:
            months_data[ym_key] = rev
            print(f" ✅ {len(rev)} 家公司")
        else:
            print(f" ⚠️ 無資料")
        time.sleep(2)

    print(f"\n📊 共取得 {len(months_data)} 個月份的營收資料")

    # 3. 計算子產業 YoY
    print("\n🔢 計算子產業合計 YoY...")
    yoy_df = compute_industry_yoy(industry_stocks, months_data)
    print(f"  ✅ 產出 {len(yoy_df)} 筆紀錄")

    latest_month = sorted(yoy_df["年月"].unique())[-1] if len(yoy_df) > 0 else "N/A"

    # 4. Top 5 動能排行
    print("\n🏆 計算 Top 5 動能排行...")
    top5 = compute_top5(yoy_df, cyclical_keywords)

    if top5:
        print(f"\n{'='*60}")
        print(f"📊 YoY 動能 Top {TOP_N} ({latest_month})")
        print(f"{'='*60}")
        for t in top5:
            emoji = RANK_EMOJI.get(t["排名"], "▪️")
            cyc_tag = " 🔄" if t["景氣循環"] else ""
            print(f"  {emoji} {t['子產業']}{cyc_tag}  ({t['成分股數']}檔)")
            print(f"     YoY: {t['最新YoY%']:+.1f}%  3M均: {t['近3月YoY均']:+.1f}%  ▲加速中")
            print(f"     趨勢: {t['近期趨勢']}")
    else:
        print("  （無符合條件的產業）")

    # 5. 偵測 負翻正 + 破百
    print("\n🔍 偵測 YoY 訊號...")
    neg_to_pos, over_100 = detect_signals(yoy_df, cyclical_keywords)

    if neg_to_pos:
        print(f"\n🔥 YoY 負翻正 ({len(neg_to_pos)} 個):")
        for n in neg_to_pos:
            cyc_tag = " 🔄" if n["景氣循環"] else ""
            print(f"  ▸ {n['子產業']}{cyc_tag}  {n['前月YoY%']:+.1f}% → {n['最新YoY%']:+.1f}%")
    else:
        print("  ✅ 無負翻正")

    if over_100:
        print(f"\n💯 YoY ≥ 100% ({len(over_100)} 個):")
        for o in over_100:
            cyc_tag = " 🔄" if o["景氣循環"] else ""
            print(f"  ▸ {o['子產業']}{cyc_tag}  YoY: {o['最新YoY%']:+.1f}%  ({o['成分股數']}檔)")
            print(f"    趨勢: {o['近期趨勢']}")
    else:
        print("  ✅ 無破百產業")

    # 6. TG 推送
    send_tg_report(top5, neg_to_pos, over_100, latest_month)

    print(f"\n{'='*60}")
    print("✅ 執行完成！")
    print(f"{'='*60}")


if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    main()
