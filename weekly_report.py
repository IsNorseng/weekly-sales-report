# Еженедельная выгрузка для коммерческого отчёта
# Код совпадает с расчётными ячейками weekly_report.ipynb.

# %% 1. Настройки
import os
from pathlib import Path
from datetime import datetime, timedelta, timezone

import numpy as np
import pandas as pd

PROJECT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
MODE = "demo"  # Для своей базы: "database". Сам по себе .env режим не меняет.
if MODE == "database":
    from dotenv import load_dotenv
    load_dotenv(PROJECT_DIR / ".env", override=False)

if MODE not in {"demo", "database"}:
    raise ValueError("MODE: demo или database")

# Конец периода не включается. По умолчанию база отдаёт последнюю полную неделю.
today_msk = datetime.now(timezone(timedelta(hours=3))).date()
default_end = pd.Timestamp(today_msk - timedelta(days=today_msk.weekday()))
if MODE == "demo":
    START, END = pd.Timestamp("2026-11-09"), pd.Timestamp("2026-11-16")
    OFFER_FILTER, GEO_FILTER = None, None
else:
    END = pd.Timestamp(os.getenv("REPORT_END") or default_end)
    START = pd.Timestamp(os.getenv("REPORT_START") or END - pd.Timedelta(days=7))
    OFFER_FILTER = os.getenv("PG_OFFER_ID") or None
    GEO_FILTER = os.getenv("PG_GEO") or None

if START.tzinfo or END.tzinfo or START != START.normalize() or END != END.normalize():
    raise ValueError("REPORT_START / REPORT_END: даты YYYY-MM-DD без времени")
if START >= END:
    raise ValueError("Начало периода должно быть раньше конца")

MIN_LEADS = 80
APPROVAL_GAP = -0.05  # На 5 процентных пунктов ниже конверсии оффера.
CHECK_RATIO = 0.80
BASE_CURRENCY = "USD"
CURRENCIES = ["EUR", "PLN", "MXN"]  # Демонстрационный набор, можно заменить.
OUTPUT = PROJECT_DIR / "report.xlsx"  # Постоянный путь для подключения к Excel-отчёту.

print(f"Режим: {MODE}. Период: {START:%d.%m.%Y} — {END - pd.Timedelta(days=1):%d.%m.%Y}")

# %% 2. Источники
def make_demo_data(start, end):
    """Независимые вымышленные выборки. Ни одной строки из рабочей базы."""
    rng = np.random.default_rng(42)
    days = pd.date_range(start, end - pd.Timedelta(days=1), freq="D")
    pg_rows, sales_rows = [], []
    for offer_no, currency in enumerate(CURRENCIES, start=1):
        offer = f"O{offer_no:02d}"
        for partner_no, (approval, check) in enumerate(
            [(0.68, 95), (0.20, 52), (0.65, 96), (0.00, 90)], start=1
        ):
            # Маленькая группа остаётся в статистике, но не в списке отклонений.
            size = 32 if partner_no == 3 else 120
            for _ in range(size):
                approved = bool(rng.random() < approval)
                sales_rows.append({
                    "lead_id": f"S{len(sales_rows) + 1:06d}",
                    "ordered_at": pd.Timestamp(rng.choice(days)) + pd.Timedelta(hours=int(rng.integers(0, 24))),
                    "partner_id": f"P{partner_no:02d}", "offer_id": offer,
                    "offer_name": f"Продукт {offer_no}", "currency": currency,
                    "approved": approved,
                    "order_amount": round(float(rng.normal(check, 6)), 2),
                })
        for day in days:
            for _ in range(int(rng.integers(22, 40))):
                pg_rows.append({
                    "lead_id": f"L{len(pg_rows) + 1:06d}",
                    "created_at": day + pd.Timedelta(hours=int(rng.integers(0, 24))),
                    "offer_id": offer, "offer_name": f"Продукт {offer_no}",
                    "geo": f"R{offer_no:02d}",
                    "status": rng.choice(["pending", "approved", "rejected", "trash"],
                                         p=[0.12, 0.48, 0.26, 0.14]),
                })
    fx_rows = []
    for i, (currency, rate) in enumerate(zip(CURRENCIES, [0.92, 4.05, 18.40]), start=1):
        fx_rows.extend([
            [i * 10, BASE_CURRENCY, currency, rate, end - pd.Timedelta(days=2)],
            [i * 10 + 1, BASE_CURRENCY, currency, rate * 1.01, end + pd.Timedelta(days=1)],
        ])
    return (pd.DataFrame(pg_rows), pd.DataFrame(sales_rows),
            pd.DataFrame(fx_rows, columns=["rate_id", "base_currency", "quote_currency", "rate", "rate_at"]))


# Имена таблиц и полей условные. Это договорённость о структуре источников,
# а не схема рабочей системы. При подключении адаптирую только эти запросы.
PG_LEADS_SQL = """
SELECT l.lead_id,
       to_timestamp(l.created_epoch) AT TIME ZONE 'Europe/Moscow' AS created_at,
       l.offer_id, o.offer_name, o.geo, l.status
FROM reporting.lead_events AS l
LEFT JOIN reporting.offers AS o ON l.offer_id = o.offer_id
WHERE l.created_epoch >= :start_epoch AND l.created_epoch < :end_epoch
  AND l.status IS DISTINCT FROM 'excluded'
"""

MSSQL_LEADS_SQL = """
SELECT l.lead_id, l.ordered_at, p.partner_id,
       o.offer_id, o.offer_name, o.currency,
       l.approved, l.order_amount
FROM reporting.sales_leads AS l
LEFT JOIN reporting.offers AS o ON l.offer_key = o.offer_key
LEFT JOIN reporting.partners AS p ON l.partner_key = p.partner_key
WHERE l.ordered_at >= :start_at AND l.ordered_at < :end_at
  AND l.lead_type = 'standard'
"""

FX_SQL = """
SELECT rate_id, base_currency, quote_currency, rate, rate_at
FROM (
    SELECT rate_id, base_currency, quote_currency, rate, rate_at,
           ROW_NUMBER() OVER (
               PARTITION BY base_currency, quote_currency
               ORDER BY rate_at DESC, rate_id DESC
           ) AS rn
    FROM reporting.fx_rates
    WHERE base_currency = :base_currency AND rate_at < :end_at
) AS latest
WHERE rn = 1
"""


def read_database_sources(start, end):
    from sqlalchemy import URL, create_engine, text

    keys = [f"{prefix}_{field}" for prefix in ("PG", "MSSQL")
            for field in ("HOST", "DATABASE", "USER", "PASSWORD")]
    missing = [key for key in keys if not os.getenv(key)]
    if missing:
        raise ValueError("Не заполнены настройки: " + ", ".join(missing))

    pg_url = URL.create("postgresql+psycopg2", username=os.environ["PG_USER"],
                        password=os.environ["PG_PASSWORD"], host=os.environ["PG_HOST"],
                        port=int(os.getenv("PG_PORT", "5432")), database=os.environ["PG_DATABASE"])
    ms_url = URL.create("mssql+pyodbc", username=os.environ["MSSQL_USER"],
                        password=os.environ["MSSQL_PASSWORD"], host=os.environ["MSSQL_HOST"],
                        port=int(os.getenv("MSSQL_PORT", "1433")), database=os.environ["MSSQL_DATABASE"],
                        query={"driver": os.getenv("MSSQL_DRIVER", "ODBC Driver 18 for SQL Server"),
                               "Encrypt": "yes", "TrustServerCertificate": "no"})
    params = {
        "start_epoch": int(start.tz_localize("Europe/Moscow").timestamp()),
        "end_epoch": int(end.tz_localize("Europe/Moscow").timestamp()),
        "start_at": start.to_pydatetime(), "end_at": end.to_pydatetime(),
        "base_currency": BASE_CURRENCY,
    }
    pg_engine = create_engine(pg_url, hide_parameters=True)
    ms_engine = create_engine(ms_url, hide_parameters=True)
    try:
        # Оба чтения PostgreSQL используют один снимок. SQL Server читается одним запросом.
        with pg_engine.connect().execution_options(isolation_level="REPEATABLE READ") as conn:
            with conn.begin():
                pg = pd.read_sql_query(text(PG_LEADS_SQL), conn, params=params)
                fx = pd.read_sql_query(text(FX_SQL), conn, params=params)
        with ms_engine.connect() as conn:
            sales = pd.read_sql_query(text(MSSQL_LEADS_SQL), conn, params=params)
        return pg, sales, fx
    finally:
        pg_engine.dispose()
        ms_engine.dispose()


if MODE == "demo":
    leads_pg, leads_sales, rates_raw = make_demo_data(START, END)
else:
    leads_pg, leads_sales, rates_raw = read_database_sources(START, END)

print(f"Получено строк: лиды PostgreSQL — {len(leads_pg):,}, SQL Server — {len(leads_sales):,}.")

# %% 3. Подготовка и проверки
def validate_leads(frame, source, required):
    if frame[required].isna().any().any():
        raise ValueError(f"{source}: пропущены обязательные поля или не найден справочник")
    if frame["lead_id"].duplicated().any():
        raise ValueError(f"{source}: повторяются lead_id; проверьте источник и соединения")


def prepare_sources(pg, sales, start, end, offer_id=None, geo=None):
    pg, sales = pg.copy(), sales.copy()
    validate_leads(pg, "PostgreSQL", ["lead_id", "created_at", "offer_id", "offer_name", "geo", "status"])
    validate_leads(sales, "SQL Server", ["lead_id", "ordered_at", "offer_id", "offer_name", "partner_id", "currency", "approved"])
    pg["created_at"] = pd.to_datetime(pg["created_at"], errors="raise")
    sales["ordered_at"] = pd.to_datetime(sales["ordered_at"], errors="raise")
    if pg["created_at"].isna().any() or sales["ordered_at"].isna().any():
        raise ValueError("После преобразования даты появились пропуски")
    if not pg["status"].isin(["pending", "approved", "rejected", "trash", "excluded"]).all():
        raise ValueError("PostgreSQL: появился неизвестный статус")
    if not sales["approved"].isin([0, 1, False, True]).all():
        raise ValueError("SQL Server: approved должен содержать 0/1")
    sales["approved"] = sales["approved"].astype(bool)
    sales["order_amount"] = pd.to_numeric(sales["order_amount"], errors="raise")
    if not np.isfinite(sales.loc[sales["approved"], "order_amount"]).all():
        raise ValueError("У одобренного заказа отсутствует или некорректен чек")
    # Один ID не должен обозначать разные продукты. Валюта входит в ключ расчёта.
    for frame in (pg, sales):
        if frame.groupby("offer_id")["offer_name"].nunique().gt(1).any():
            raise ValueError("У одного offer_id несколько названий")
    pg = pg.loc[pg["created_at"].ge(start) & pg["created_at"].lt(end) & pg["status"].ne("excluded")].copy()
    sales = sales.loc[sales["ordered_at"].ge(start) & sales["ordered_at"].lt(end)].copy()
    if offer_id is not None:
        pg = pg.loc[pg["offer_id"].astype(str).eq(str(offer_id))].copy()
    if geo is not None:
        pg = pg.loc[pg["geo"].eq(geo)].copy()
    if pg.empty or sales.empty:
        raise ValueError("Одна из выборок пуста. Проверьте период и фильтры; файл не обновлён")
    return pg, sales


leads_pg, leads_sales = prepare_sources(leads_pg, leads_sales, START, END, OFFER_FILTER, GEO_FILTER)
print("Дубликатов ID и пропусков в ключах нет. Период и статусы проверены.")

# %% 4. Метрики
def divide(numerator, denominator):
    # Нулевой числитель — это 0%. Нулевой знаменатель — нет значения.
    return numerator.div(denominator.where(denominator.ne(0)))


def daily_metrics(leads, start, end):
    data = leads.assign(
        date=leads["created_at"].dt.normalize(),
        clean=leads["status"].isin(["pending", "approved", "rejected"]),
        approved=leads["status"].eq("approved"),
        pending=leads["status"].eq("pending"),
        rejected=leads["status"].eq("rejected"),
        trash=leads["status"].eq("trash"),
    )
    result = data.groupby("date").agg(
        total_leads=("lead_id", "size"), clean_leads=("clean", "sum"),
        approved_leads=("approved", "sum"), pending_leads=("pending", "sum"),
        rejected_leads=("rejected", "sum"), trash_leads=("trash", "sum"),
    ).reindex(pd.date_range(start, end - pd.Timedelta(days=1)), fill_value=0).rename_axis("date").reset_index()
    result["approval_clean"] = divide(result["approved_leads"], result["clean_leads"])
    result["approval_total"] = divide(result["approved_leads"], result["total_leads"])
    return result


def partner_metrics(sales):
    data = sales.assign(approved_amount=sales["order_amount"].where(sales["approved"]))
    keys = ["offer_id", "offer_name", "currency"]
    offers = data.groupby(keys, as_index=False).agg(
        total_leads=("lead_id", "size"), approved_leads=("approved", "sum"),
        avg_check=("approved_amount", "mean"), active_partners=("partner_id", "nunique"),
    )
    offers["approval_total"] = divide(offers["approved_leads"], offers["total_leads"])
    partners = data.groupby(keys + ["partner_id"], as_index=False).agg(
        total_leads=("lead_id", "size"), approved_leads=("approved", "sum"),
        avg_check=("approved_amount", "mean"),
    )
    partners["approval_total"] = divide(partners["approved_leads"], partners["total_leads"])
    baseline = offers[keys + ["approval_total", "avg_check"]].rename(columns={
        "approval_total": "offer_approval", "avg_check": "offer_avg_check"})
    partners = partners.merge(baseline, on=keys, how="left", validate="many_to_one")
    partners["approval_gap"] = partners["approval_total"] - partners["offer_approval"]
    partners["check_gap"] = divide(partners["avg_check"], partners["offer_avg_check"]) - 1
    low_approval = partners["approval_gap"].lt(APPROVAL_GAP) & ~np.isclose(
        partners["approval_gap"], APPROVAL_GAP, rtol=0, atol=1e-12)
    low_check = partners["check_gap"].lt(CHECK_RATIO - 1) & ~np.isclose(
        partners["check_gap"], CHECK_RATIO - 1, rtol=0, atol=1e-12)
    partners["reason"] = np.select(
        [low_approval & low_check, low_approval, low_check],
        ["Конверсия и чек", "Конверсия", "Средний чек"], default="",
    )
    partners["enough_leads"] = partners["total_leads"].ge(MIN_LEADS)
    partners = partners.sort_values(["offer_id", "currency", "partner_id"]).reset_index(drop=True)
    anomalies = partners.loc[partners["enough_leads"] & partners["reason"].ne("")].copy()
    anomalies = anomalies.sort_values("total_leads", ascending=False, kind="stable").reset_index(drop=True)
    return offers, partners, anomalies


def currency_snapshot(raw, end):
    rates = raw.copy()
    required = ["rate_id", "base_currency", "quote_currency", "rate", "rate_at"]
    if rates[required].isna().any().any():
        raise ValueError("В курсах есть пропуски")
    rates["rate_at"] = pd.to_datetime(rates["rate_at"], errors="raise")
    if rates["rate_at"].isna().any():
        raise ValueError("В датах курсов появились пропуски")
    rates = rates.loc[rates["rate_at"].lt(end) & rates["base_currency"].eq(BASE_CURRENCY)
                      & rates["quote_currency"].isin(CURRENCIES)].copy()
    if rates["rate_id"].duplicated().any():
        raise ValueError("В курсах повторяется rate_id")
    rates = rates.sort_values(["rate_at", "rate_id"]).drop_duplicates(
        ["base_currency", "quote_currency"], keep="last")
    rates["rate"] = pd.to_numeric(rates["rate"], errors="raise")
    if not (np.isfinite(rates["rate"]) & rates["rate"].gt(0)).all():
        raise ValueError("Курс должен быть конечным положительным числом")
    if set(CURRENCIES) - set(rates["quote_currency"]):
        raise ValueError("Для части выбранных валют нет курса до конца периода")
    rates["inverse_rate"] = 1 / rates["rate"]
    rates["age_days"] = (end.normalize() - rates["rate_at"].dt.normalize()).dt.days
    return rates[["base_currency", "quote_currency", "rate", "inverse_rate", "rate_at", "age_days"]].sort_values("quote_currency").reset_index(drop=True)


daily = daily_metrics(leads_pg, START, END)
offers, partners, anomalies = partner_metrics(leads_sales)
rates = currency_snapshot(rates_raw, END)
print(f"Групп партнёр–оффер–валюта: {len(partners)}. Для проверки: {len(anomalies)}.")
daily

# %% 5. Excel
# Английские имена полей и Excel-таблиц стабильны между запусками.
# Числа остаются числами, даты — датами, доли хранятся как 0..1.
def export_excel(tables, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.stem + ".tmp.xlsx")
    try:
        with pd.ExcelWriter(temp, engine="xlsxwriter", datetime_format="yyyy-mm-dd",
                            engine_kwargs={"options": {"strings_to_formulas": False,
                                                       "strings_to_urls": False}}) as writer:
            book = writer.book
            header = book.add_format({"bold": True, "font_color": "white", "bg_color": "#326B78",
                                      "text_wrap": True, "valign": "vcenter", "align": "center"})
            integer = book.add_format({"num_format": "#,##0"})
            money = book.add_format({"num_format": "#,##0.00"})
            percent = book.add_format({"num_format": "0.0%"})
            fx_format = book.add_format({"num_format": "0.000000"})
            date_format = book.add_format({"num_format": "yyyy-mm-dd"})
            wrapped = book.add_format({"text_wrap": True, "valign": "vcenter"})
            alert = book.add_format({"font_color": "#9C3B20", "bg_color": "#FBECE6"})
            for sheet_name, (table_name, frame) in tables.items():
                if len(frame) > 1_048_575:
                    raise ValueError(f"{sheet_name}: превышен лимит строк Excel")
                frame.to_excel(writer, sheet_name=sheet_name, index=False)
                sheet = writer.sheets[sheet_name]
                sheet.hide_gridlines(2)
                sheet.freeze_panes(1, 0)
                sheet.set_row(0, 42)
                # Пустой список отклонений сохраняет заголовки и имя таблицы.
                last_row = max(1, len(frame))
                sheet.add_table(0, 0, last_row, len(frame.columns) - 1, {
                    "name": table_name, "style": "Table Style Medium 2",
                    "columns": [{"header": col, "header_format": header} for col in frame.columns],
                })
                for col_no, col in enumerate(frame.columns):
                    fmt = None
                    if col in {"date", "rate_at"}:
                        fmt = date_format
                    elif col in {"approval_clean", "approval_total", "offer_approval", "approval_gap", "check_gap"}:
                        fmt = percent
                    elif col in {"avg_check", "offer_avg_check"}:
                        fmt = money
                    elif col in {"rate", "inverse_rate"}:
                        fmt = fx_format
                    elif pd.api.types.is_numeric_dtype(frame[col]) and not pd.api.types.is_bool_dtype(frame[col]):
                        fmt = integer
                    longest = max([len(str(col))] + [len(str(v)) for v in frame[col].head(500).dropna()])
                    width = min(48, max(15, longest + 2))
                    if col == "value":
                        width = 95
                        fmt = wrapped
                    sheet.set_column(col_no, col_no, width, fmt)
                if sheet_name == "Параметры":
                    for row_no, value in enumerate(frame["value"], start=1):
                        sheet.set_row(row_no, 30 if len(str(value)) > 90 else 20)
                if sheet_name == "Отклонения" and len(frame):
                    col_no = frame.columns.get_loc("reason")
                    sheet.conditional_format(1, col_no, len(frame), col_no,
                                             {"type": "no_blanks", "format": alert})
                sheet.set_landscape()
                sheet.fit_to_pages(1, 0)
                sheet.repeat_rows(0)
        # Если расчёт или сохранение оборвётся, предыдущая выгрузка сохранится.
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


parameters = pd.DataFrame([
    ("mode", "Демонстрационные данные" if MODE == "demo" else "Данные из подключённых баз"),
    ("period_start", START.strftime("%Y-%m-%d")),
    ("period_end_exclusive", END.strftime("%Y-%m-%d")),
    ("timezone", "Europe/Moscow; SQL Server и fx_rates ожидаются в местном времени"),
    ("pg_filter", f"offer_id={OFFER_FILTER or 'все'}; geo={GEO_FILTER or 'все'}"),
    ("pg_leads", len(leads_pg)), ("sales_leads", len(leads_sales)),
    ("minimum_leads", MIN_LEADS), ("approval_gap_threshold", APPROVAL_GAP),
    ("check_ratio_threshold", CHECK_RATIO),
    ("approval_clean", "approved / (pending + approved + rejected)"),
    ("approval_total", "approved / total; excluded исключён в PostgreSQL"),
    ("avg_check", "Средний чек одобренных заказов в валюте строки"),
    ("offer_baseline", "Среднее по офферу и валюте, включая проверяемого партнёра"),
    ("active_partners", "Все партнёры с лидами, включая партнёров без одобрений"),
    ("missing_values", "Пустая ячейка: метрика не определена. 0: рассчитанный ноль"),
    ("sources", "По_дням: PostgreSQL. По_офферам, По_партнёрам, Отклонения: SQL Server"),
    ("currency_rate", f"rate: единиц quote_currency за 1 {BASE_CURRENCY}; inverse_rate: обратный курс"),
    ("currency_date", "Последний курс строго до конца периода; age_days показывает его возраст"),
    ("currency_use", "Курсы выгружаются отдельно. Пересчёт чеков и выручки не выполняется"),
    ("snapshot", "Статусы текущие на момент чтения. Между двумя базами нет общего снимка"),
    ("built_at_moscow", datetime.now(timezone(timedelta(hours=3))).strftime("%Y-%m-%d %H:%M:%S")),
], columns=["parameter", "value"])

tables = {
    "Параметры": ("RunInfo", parameters),
    "По_дням": ("DailyMetrics", daily),
    "По_офферам": ("OfferMetrics", offers),
    "По_партнёрам": ("PartnerMetrics", partners),
    "Отклонения": ("Anomalies", anomalies),
    "Курсы": ("CurrencyRates", rates),
}
export_excel(tables, OUTPUT)
print(f"Сохранено: {OUTPUT.name}. Листов: {len(tables)}.")
pd.DataFrame({"Лист": list(tables), "Строк": [len(frame) for _, frame in tables.values()]})
