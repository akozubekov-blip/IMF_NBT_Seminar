from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# =================================================
# Streamlit page configuration
# =================================================
st.set_page_config(
    page_title="IMF Workshop for the NBT - FX Market Development",
    page_icon=":bar_chart:",
    layout="wide",
)

# =================================================
# File paths
# =================================================
NONCASH_FILE = "bank_oper_full_nc.xlsx"
CASH_FILE = "bank_oper_full_c.xlsx"
BANK_CODES_FILE = "bank_codes.xlsx"

# Use curr_codes.xls if it exists, otherwise use curr_codes.xlsx
if Path("curr_codes.xls").exists():
    CURR_CODES_FILE = "curr_codes.xls"
else:
    CURR_CODES_FILE = "curr_codes.xlsx"

# =================================================
# Helper functions
# =================================================
def normalize_code_column(series):
    """
    Cleans Excel code columns.

    Examples:
    840.0 -> 840
    " USD " -> USD
    """
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

def move_option_to_front(options, preferred_option="USD"):
    options = list(options)

    matched_option = next(
        (
            option
            for option in options
            if str(option).strip().upper() == preferred_option.upper()
        ),
        None,
    )

    if matched_option is None:
        return options

    return [matched_option] + [
        option for option in options if option != matched_option
    ]

def prepare_date_column(df, date_column="DATVAL"):
    df[date_column] = pd.to_datetime(df[date_column], errors="coerce")
    return df.dropna(subset=[date_column])

def prepare_numeric_columns(df, columns):
    for column in columns:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    return df

def get_sidebar_date_range(df, key_prefix):
    min_date = df["DATVAL"].min().date()
    max_date = df["DATVAL"].max().date()

    start_date = st.sidebar.date_input(
        "Выберите начальную дату:",
        min_value=min_date,
        max_value=max_date,
        value=min_date,
        key=f"{key_prefix}_start_date",
    )

    end_date = st.sidebar.date_input(
        "Выберите конечную дату:",
        min_value=min_date,
        max_value=max_date,
        value=max_date,
        key=f"{key_prefix}_end_date",
    )

    if start_date > end_date:
        st.sidebar.error("Начальная дата не может быть позже конечной даты.")
        st.stop()

    return pd.Timestamp(start_date), pd.Timestamp(end_date)


def filter_operations(df, banks=None, currency=None, start_date=None, end_date=None):
    mask = pd.Series(True, index=df.index)

    if banks is not None:
        mask &= df["BANK_NAME"].isin(banks)

    if currency is not None:
        mask &= df["VAL_NAME"].eq(currency)

    if start_date is not None and end_date is not None:
        mask &= df["DATVAL"].between(start_date, end_date)

    return df.loc[mask].copy()


def group_sum(df, group_columns, value_column):
    return (
        df.groupby(group_columns, as_index=False)[value_column]
        .sum()
        .sort_values(group_columns)
    )


def plot_chart(df, x_col, y_col, title, color_col="BANK_NAME"):
    color = color_col if color_col in df.columns else None

    fig = px.bar(
        df,
        x=x_col,
        y=y_col,
        color=color,
        orientation="v",
        title=title,
        template="plotly_white",
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False),
    )

    return fig


# =================================================
# Data loading
# =================================================
@st.cache_data
def load_noncash_data_from_excel():
    return pd.read_excel(
        io=NONCASH_FILE,
        engine="openpyxl",
        sheet_name="noncash_oper",
        skiprows=2,
    )


@st.cache_data
def load_cash_data_from_excel():
    return pd.read_excel(
        io=CASH_FILE,
        engine="openpyxl",
        sheet_name="cash_oper",
        skiprows=2,
    )


@st.cache_data
def load_bank_names():
    return pd.read_excel(
        io=BANK_CODES_FILE,
        engine="openpyxl",
        skiprows=0,
    )

@st.cache_data
def load_currency_names():
    # For .xls files, pandas may require:
    # pip install xlrd
    return pd.read_excel(
        io=CURR_CODES_FILE,
        skiprows=0,
    )


# =================================================
# Load source data
# =================================================
df_noncash_raw = load_noncash_data_from_excel()
df_cash_raw = load_cash_data_from_excel()
df_bank_names = load_bank_names()
df_curr_names = load_currency_names()


# =================================================
# Normalize codes before merging
# =================================================
df_noncash_raw["BANK"] = normalize_code_column(df_noncash_raw["BANK"])
df_cash_raw["BANK"] = normalize_code_column(df_cash_raw["BANK"])

df_noncash_raw["VAL"] = normalize_code_column(df_noncash_raw["VAL"])
df_cash_raw["VAL"] = normalize_code_column(df_cash_raw["VAL"])

df_bank_names["BANK_CODE"] = normalize_code_column(df_bank_names["BANK_CODE"])

df_curr_names["VAL_CODE"] = normalize_code_column(df_curr_names["VAL_CODE"])
df_curr_names["VAL_NAME"] = df_curr_names["VAL_NAME"].astype("string").str.strip()

df_curr_names = df_curr_names.dropna(subset=["VAL_CODE", "VAL_NAME"])
df_curr_names = df_curr_names.drop_duplicates(subset=["VAL_CODE"])


# =================================================
# Keep only currencies listed in curr_codes.xls/xlsx
# =================================================
valid_currency_codes = set(df_curr_names["VAL_CODE"])

df_noncash_raw = df_noncash_raw[
    df_noncash_raw["VAL"].isin(valid_currency_codes)
].copy()

df_cash_raw = df_cash_raw[
    df_cash_raw["VAL"].isin(valid_currency_codes)
].copy()


# =================================================
# Prepare dates and numeric columns
# =================================================
df_noncash_raw = prepare_date_column(df_noncash_raw)
df_cash_raw = prepare_date_column(df_cash_raw)

# Make sure required non-cash columns exist
for column in ["KONTRAKT", "EXPORT", "POS_SVOP1", "POS_SVOP2", "KUR_POK", "KUR_PR"]:
    if column not in df_noncash_raw.columns:
        df_noncash_raw[column] = 0

# Make sure required cash columns exist
for column in ["KUR_POK", "KUR_PR"]:
    if column not in df_cash_raw.columns:
        df_cash_raw[column] = 0

df_noncash_raw = prepare_numeric_columns(
    df_noncash_raw,
    [
        "POKUPK",
        "PRODANO",
        "KONTRAKT",
        "EXPORT",
        "POS_SVOP1",
        "POS_SVOP2",
        "KUR_POK",
        "KUR_PR",
    ],
)

df_cash_raw = prepare_numeric_columns(
    df_cash_raw,
    [
        "POKUPK",
        "PRODANO",
        "KUR_POK",
        "KUR_PR",
    ],
)


# =================================================
# Merge with bank and currency names
# =================================================
df_noncash = df_noncash_raw.merge(
    df_bank_names,
    left_on="BANK",
    right_on="BANK_CODE",
    how="left",
)

df_noncash = df_noncash.merge(
    df_curr_names,
    left_on="VAL",
    right_on="VAL_CODE",
    how="left",
)

df_cash = df_cash_raw.merge(
    df_bank_names,
    left_on="BANK",
    right_on="BANK_CODE",
    how="left",
)

df_cash = df_cash.merge(
    df_curr_names,
    left_on="VAL",
    right_on="VAL_CODE",
    how="left",
)


# Fallback only for bank names.
# Do NOT use fallback for currencies, otherwise numbers like 0 or 344 may appear.
df_noncash["BANK_NAME"] = df_noncash["BANK_NAME"].fillna(df_noncash["BANK"])
df_cash["BANK_NAME"] = df_cash["BANK_NAME"].fillna(df_cash["BANK"])


# =================================================
# Stop app if filtered data is empty
# =================================================
if df_noncash.empty and df_cash.empty:
    st.error(
        "После фильтрации не осталось данных. "
        "Проверьте, совпадают ли коды валют в операционных файлах и curr_codes.xls/xlsx."
    )
    st.stop()


# =================================================
# Sidebar
# =================================================
st.sidebar.header("Анализ валютного рынка")

structure_option = st.sidebar.radio(
    "Выберите структуру отображения:",
    (
        "Динамика обменных курсов",
        "Валютный рынок в разрезе Банков",
        "Индикаторы валютного рынка",
        "Обслуживание внешней торговли",
    ),
)


bank_options = sorted(
    pd.concat(
        [
            df_noncash["BANK_NAME"],
            df_cash["BANK_NAME"],
        ]
    )
    .dropna()
    .unique()
)


# Currency sidebar options ONLY from curr_codes.xls/xlsx
currency_options = sorted(
    pd.concat(
        [
            df_noncash["VAL_NAME"],
            df_cash["VAL_NAME"],
        ]
    )
    .dropna()
    .unique()
)

currency_options = move_option_to_front(currency_options, "USD")

if not currency_options:
    st.error(
        "Не найдены валюты для отображения. "
        "Проверьте файл curr_codes.xls/xlsx и соответствие кодов валют."
    )
    st.stop()

# =================================================
# 0. Exchange rate dynamics
# =================================================
if structure_option == "Динамика обменных курсов":

    st.title(":chart_with_upwards_trend: Динамика обменных курсов")

    # -------------------------------------------------
    # Helper functions for exchange-rate dynamics
    # -------------------------------------------------
    def build_rate_base(df, operation_type):
        rate_df = df[
            [
                "DATVAL",
                "BANK_NAME",
                "VAL_NAME",
                "POKUPK",
                "PRODANO",
                "KUR_POK",
                "KUR_PR",
            ]
        ].copy()

        rate_df["OPERATION_TYPE"] = operation_type

        rate_df["DATVAL"] = pd.to_datetime(
            rate_df["DATVAL"],
            errors="coerce",
        )

        for column in ["POKUPK", "PRODANO", "KUR_POK", "KUR_PR"]:
            rate_df[column] = pd.to_numeric(
                rate_df[column],
                errors="coerce",
            ).fillna(0)

        rate_df = rate_df.dropna(
            subset=[
                "DATVAL",
                "BANK_NAME",
                "VAL_NAME",
            ]
        )

        rate_df["DATVAL_DAY"] = rate_df["DATVAL"].dt.normalize()

        return rate_df


    def weighted_average_rate(group, rate_column, weight_column):
        rates = pd.to_numeric(
            group[rate_column],
            errors="coerce",
        )

        weights = pd.to_numeric(
            group[weight_column],
            errors="coerce",
        ).fillna(0).abs()

        valid_mask = (
            rates.notna()
            & weights.notna()
            & (weights > 0)
            & (rates > 0)
        )

        if valid_mask.any():
            return np.average(
                rates.loc[valid_mask],
                weights=weights.loc[valid_mask],
            )

        valid_rates = rates[(rates.notna()) & (rates > 0)]

        if not valid_rates.empty:
            return valid_rates.mean()

        return np.nan


    def calculate_daily_weighted_rates(df):
        group_columns = [
            "DATVAL_DAY",
            "VAL_NAME",
        ]

        rate_rows = []

        for group_values, group_df in df.groupby(group_columns):
            datval_day, val_name = group_values

            rate_rows.append(
                {
                    "DATVAL": datval_day,
                    "VAL_NAME": val_name,
                    "KUR_POK_WEIGHTED": weighted_average_rate(
                        group_df,
                        "KUR_POK",
                        "POKUPK",
                    ),
                    "KUR_PR_WEIGHTED": weighted_average_rate(
                        group_df,
                        "KUR_PR",
                        "PRODANO",
                    ),
                    "POKUPK_TOTAL": group_df["POKUPK"].sum(),
                    "PRODANO_TOTAL": group_df["PRODANO"].sum(),
                }
            )

        return pd.DataFrame(rate_rows).sort_values(
            [
                "VAL_NAME",
                "DATVAL",
            ]
        )


    def render_rate_section(df_source, section_title):
        if df_source.empty:
            st.warning(f"Нет данных для раздела: {section_title}")
            return

        df_rates = calculate_daily_weighted_rates(df_source)

        if df_rates.empty:
            st.warning(f"Не удалось рассчитать курсы для раздела: {section_title}")
            return

        st.subheader(section_title)

        currencies = sorted(
            df_rates["VAL_NAME"]
            .dropna()
            .unique()
        )

        currencies = move_option_to_front(
            currencies,
            "USD",
        )

        if not currencies:
            st.warning(f"Нет валют для отображения в разделе: {section_title}")
            return

        # Show two currency charts per row
        for row_start in range(0, len(currencies), 2):
            row_currencies = currencies[row_start:row_start + 2]
            chart_cols = st.columns(2)

            for chart_col, currency in zip(chart_cols, row_currencies):
                with chart_col:
                    df_currency_rates = df_rates[
                        df_rates["VAL_NAME"] == currency
                        ].copy()

                    df_currency_rates_long = df_currency_rates.melt(
                        id_vars=["DATVAL", "VAL_NAME"],
                        value_vars=[
                            "KUR_POK_WEIGHTED",
                            "KUR_PR_WEIGHTED",
                        ],
                        var_name="RATE_TYPE",
                        value_name="RATE_VALUE",
                    )

                    df_currency_rates_long["RATE_TYPE"] = (
                        df_currency_rates_long["RATE_TYPE"]
                        .replace(
                            {
                                "KUR_POK_WEIGHTED": "Средневзвешенный курс покупки",
                                "KUR_PR_WEIGHTED": "Средневзвешенный курс продажи",
                            }
                        )
                    )

                    df_currency_rates_long = df_currency_rates_long.dropna(
                        subset=["RATE_VALUE"]
                    )

                    df_currency_rates_long = df_currency_rates_long[
                        df_currency_rates_long["RATE_VALUE"] > 0
                        ]

                    if df_currency_rates_long.empty:
                        st.info(f"Нет данных по курсам для валюты: {currency}")
                        continue

                    fig_currency_rates = px.line(
                        df_currency_rates_long,
                        x="DATVAL",
                        y="RATE_VALUE",
                        color="RATE_TYPE",
                        markers=True,
                        title=f"<b>{section_title}: {currency}</b>",
                        template="plotly_white",
                    )

                    fig_currency_rates.update_layout(
                        plot_bgcolor="rgba(0,0,0,0)",
                        xaxis=dict(showgrid=False),
                        yaxis_title="Обменный курс",
                        legend_title_text="Тип курса",
                    )

                    st.plotly_chart(
                        fig_currency_rates,
                        use_container_width=True,
                    )


    # -------------------------------------------------
    # Build separate NonCash and Cash rate datasets
    # -------------------------------------------------
    df_noncash_rates = build_rate_base(
        df_noncash,
        "Безналичные операции",
    )

    df_cash_rates = build_rate_base(
        df_cash,
        "Наличные операции",
    )

    df_rates_base = pd.concat(
        [
            df_noncash_rates,
            df_cash_rates,
        ],
        ignore_index=True,
    )

    if df_rates_base.empty:
        st.warning("Нет данных для отображения динамики обменных курсов.")
        st.stop()

    # -------------------------------------------------
    # Sidebar filters
    # -------------------------------------------------
    min_date = df_rates_base["DATVAL"].min().date()
    max_date = df_rates_base["DATVAL"].max().date()

    start_date = st.sidebar.date_input(
        "Выберите начальную дату:",
        min_value=min_date,
        max_value=max_date,
        value=min_date,
        key="fx_rates_start_date",
    )

    end_date = st.sidebar.date_input(
        "Выберите конечную дату:",
        min_value=min_date,
        max_value=max_date,
        value=max_date,
        key="fx_rates_end_date",
    )

    if start_date > end_date:
        st.sidebar.error("Начальная дата не может быть позже конечной даты.")
        st.stop()

    rate_bank_options = sorted(
        df_rates_base["BANK_NAME"].dropna().unique()
    )

    rate_currency_options = sorted(
        df_rates_base["VAL_NAME"].dropna().unique()
    )

    rate_currency_options = move_option_to_front(
        rate_currency_options,
        "USD",
    )

    default_rate_currencies = (
        ["USD"]
        if "USD" in rate_currency_options
        else rate_currency_options[:1]
    )

    selected_rate_banks = st.sidebar.multiselect(
        "Выберите Банк:",
        options=rate_bank_options,
        default=rate_bank_options,
        key="fx_rates_banks",
    )

    selected_rate_currencies = st.sidebar.multiselect(
        "Выберите Валюту:",
        options=rate_currency_options,
        default=default_rate_currencies,
        key="fx_rates_currencies",
    )

    if not selected_rate_banks:
        st.warning("Выберите хотя бы один банк.")
        st.stop()

    if not selected_rate_currencies:
        st.warning("Выберите хотя бы одну валюту.")
        st.stop()

    start_date = pd.Timestamp(start_date)
    end_date = pd.Timestamp(end_date)

    def apply_rate_filters(df):
        return df[
            df["DATVAL"].between(start_date, end_date)
            & df["BANK_NAME"].isin(selected_rate_banks)
            & df["VAL_NAME"].isin(selected_rate_currencies)
        ].copy()


    df_noncash_rates_filtered = apply_rate_filters(
        df_noncash_rates
    )

    df_cash_rates_filtered = apply_rate_filters(
        df_cash_rates
    )

    # -------------------------------------------------
    # Render separate tabs for NonCash and Cash
    # -------------------------------------------------
    noncash_rates_tab, cash_rates_tab = st.tabs(
        [
            "Безналичные операции",
            "Наличные операции",
        ]
    )

    with noncash_rates_tab:
        render_rate_section(
            df_source=df_noncash_rates_filtered,
            section_title="Безналичные операции",
        )

    with cash_rates_tab:
        render_rate_section(
            df_source=df_cash_rates_filtered,
            section_title="Наличные операции",
        )


# =================================================
# 1. External trade service / KONTRAKT and EXPORT
# =================================================
elif structure_option == "Обслуживание внешней торговли":
    banks = st.sidebar.multiselect(
        "Выберите Банк:",
        options=bank_options,
        default=bank_options,
        key="trade_banks",
    )

    if not banks:
        st.warning("Выберите хотя бы один банк.")
        st.stop()

    currency_trade = st.sidebar.selectbox(
        "Валюта контракта:",
        options=currency_options,
        key="trade_currency",
    )

    start_date, end_date = get_sidebar_date_range(
        df_noncash,
        key_prefix="trade_service",
    )

    df_selection_noncash_trade = filter_operations(
        df_noncash,
        banks=banks,
        currency=currency_trade,
        start_date=start_date,
        end_date=end_date,
    )

    kontrakt_data = group_sum(
        df_selection_noncash_trade,
        ["DATVAL", "BANK_NAME"],
        "KONTRAKT",
    )

    export_data = group_sum(
        df_selection_noncash_trade,
        ["DATVAL", "BANK_NAME"],
        "EXPORT",
    )

    st.title("Обслуживание внешней торговли")

    st.plotly_chart(
        plot_chart(
            kontrakt_data,
            "DATVAL",
            "KONTRAKT",
            "Оплата контрактов клиентов по импорту товаров и услуг",
        ),
        use_container_width=True,
    )

    st.plotly_chart(
        plot_chart(
            export_data,
            "DATVAL",
            "EXPORT",
            "Получение валюты от контрактов клиентов по экспорту товаров и услуг",
        ),
        use_container_width=True,
    )

# =================================================
# 2. FX operations by banks
# =================================================
elif structure_option == "Валютный рынок в разрезе Банков":
    banks = st.sidebar.multiselect(
        "Выберите Банк:",
        options=bank_options,
        default=bank_options,
        key="bank_structure_banks",
    )

    if not banks:
        st.warning("Выберите хотя бы один банк.")
        st.stop()

    currency_pur = st.sidebar.selectbox(
        "Покупка валюты - выберите валюту:",
        options=currency_options,
        key="bank_currency_pur",
    )

    currency_sell = st.sidebar.selectbox(
        "Продажа валюты - выберите валюту:",
        options=currency_options,
        key="bank_currency_sell",
    )

    start_date, end_date = get_sidebar_date_range(
        df_noncash,
        key_prefix="bank_structure",
    )

    df_selection_noncash_pur = filter_operations(
        df_noncash,
        banks=banks,
        currency=currency_pur,
        start_date=start_date,
        end_date=end_date,
    )

    df_selection_noncash_sell = filter_operations(
        df_noncash,
        banks=banks,
        currency=currency_sell,
        start_date=start_date,
        end_date=end_date,
    )

    df_selection_cash_pur = filter_operations(
        df_cash,
        banks=banks,
        currency=currency_pur,
        start_date=start_date,
        end_date=end_date,
    )

    df_selection_cash_sell = filter_operations(
        df_cash,
        banks=banks,
        currency=currency_sell,
        start_date=start_date,
        end_date=end_date,
    )

    st.title(":bank: Операции с иностранной валютой — Банковская структура")

    st.header("Безналичные операции")
    non_cash_col1, non_cash_col2 = st.columns(2)

    with non_cash_col1:
        noncash_pur_data = group_sum(
            df_selection_noncash_pur,
            ["DATVAL", "BANK_NAME"],
            "POKUPK",
        )

        st.plotly_chart(
            plot_chart(
                noncash_pur_data,
                "DATVAL",
                "POKUPK",
                f"Покупка безналичной валюты — {currency_pur}",
            ),
            use_container_width=True,
        )

    with non_cash_col2:
        noncash_sell_data = group_sum(
            df_selection_noncash_sell,
            ["DATVAL", "BANK_NAME"],
            "PRODANO",
        )

        st.plotly_chart(
            plot_chart(
                noncash_sell_data,
                "DATVAL",
                "PRODANO",
                f"Продажа безналичной валюты — {currency_sell}",
            ),
            use_container_width=True,
        )

    st.header("Наличные операции")
    cash_col1, cash_col2 = st.columns(2)

    with cash_col1:
        cash_pur_data = group_sum(
            df_selection_cash_pur,
            ["DATVAL", "BANK_NAME"],
            "POKUPK",
        )

        st.plotly_chart(
            plot_chart(
                cash_pur_data,
                "DATVAL",
                "POKUPK",
                f"Покупка наличной валюты — {currency_pur}",
            ),
            use_container_width=True,
        )

    with cash_col2:
        cash_sell_data = group_sum(
            df_selection_cash_sell,
            ["DATVAL", "BANK_NAME"],
            "PRODANO",
        )

        st.plotly_chart(
            plot_chart(
                cash_sell_data,
                "DATVAL",
                "PRODANO",
                f"Продажа наличной валюты — {currency_sell}",
            ),
            use_container_width=True,
        )

# =================================================
# 3. FX market indicators:
#    Daily turnover, 7-day rolling volatility,
#    and swap operations
#    separated for NonCash and Cash
# =================================================
else:

    st.title(":currency_exchange: Индикаторы валютного рынка")

    ROLLING_VOL_WINDOW = 7

    # -------------------------------------------------
    # Helper functions for this section
    # -------------------------------------------------
    def build_indicator_base(df, operation_type):
        indicator_df = df[
            [
                "DATVAL",
                "BANK_NAME",
                "VAL_NAME",
                "POKUPK",
                "PRODANO",
            ]
        ].copy()

        indicator_df["OPERATION_TYPE"] = operation_type

        indicator_df["DATVAL"] = pd.to_datetime(
            indicator_df["DATVAL"],
            errors="coerce",
        )

        indicator_df["POKUPK"] = pd.to_numeric(
            indicator_df["POKUPK"],
            errors="coerce",
        ).fillna(0)

        indicator_df["PRODANO"] = pd.to_numeric(
            indicator_df["PRODANO"],
            errors="coerce",
        ).fillna(0)

        # Exchange-rate columns for bid-ask spread calculation.
        for rate_column in ["KUR_POK", "KUR_PR"]:
            if rate_column in df.columns:
                indicator_df[rate_column] = pd.to_numeric(
                    df[rate_column],
                    errors="coerce",
                ).fillna(0)
            else:
                indicator_df[rate_column] = 0

        # Swap columns exist only in non-cash data.
        # For cash data they are added as zero so the function stays universal.
        for swap_column in ["POS_SVOP1", "POS_SVOP2"]:
            if swap_column in df.columns:
                indicator_df[swap_column] = pd.to_numeric(
                    df[swap_column],
                    errors="coerce",
                ).fillna(0)
            else:
                indicator_df[swap_column] = 0

        indicator_df = indicator_df.dropna(
            subset=[
                "DATVAL",
                "BANK_NAME",
                "VAL_NAME",
            ]
        )

        indicator_df["DATVAL_DAY"] = indicator_df["DATVAL"].dt.normalize()

        return indicator_df


    def calculate_daily_fx_indicators(df, rolling_window=7):
        group_columns = [
            "DATVAL_DAY",
            "BANK_NAME",
            "VAL_NAME",
        ]

        indicators = (
            df.groupby(group_columns, as_index=False)
            .agg(
                POKUPK_DAILY_TURNOVER=("POKUPK", "sum"),
                PRODANO_DAILY_TURNOVER=("PRODANO", "sum"),
                POS_SVOP1_DAILY_TURNOVER=("POS_SVOP1", "sum"),
                POS_SVOP2_DAILY_TURNOVER=("POS_SVOP2", "sum"),

                # Number of purchase/sale transaction records
                POKUPK_TRANSACTION_COUNT=("POKUPK", lambda x: (x.abs() > 0).sum()),
                PRODANO_TRANSACTION_COUNT=("PRODANO", lambda x: (x.abs() > 0).sum()),

                OBSERVATIONS=("POKUPK", "count"),
            )
        )

        indicators["TOTAL_DAILY_TURNOVER"] = (
                indicators["POKUPK_DAILY_TURNOVER"]
                + indicators["PRODANO_DAILY_TURNOVER"]
        )

        indicators["TRANSACTION_COUNT"] = (
                indicators["POKUPK_TRANSACTION_COUNT"]
                + indicators["PRODANO_TRANSACTION_COUNT"]
        )

        indicators["AVG_TRANSACTION_SIZE"] = np.where(
            indicators["TRANSACTION_COUNT"] > 0,
            indicators["TOTAL_DAILY_TURNOVER"] / indicators["TRANSACTION_COUNT"],
            0,
        )

        indicators = indicators.rename(
            columns={
                "DATVAL_DAY": "DATVAL",
            }
        )

        indicators = indicators.sort_values(
            [
                "BANK_NAME",
                "VAL_NAME",
                "DATVAL",
            ]
        )

        def add_rolling_volatility(group):
            group = group.sort_values("DATVAL").copy()

            pokupk_rolling_mean = (
                group["POKUPK_DAILY_TURNOVER"]
                .rolling(
                    window=rolling_window,
                    min_periods=2,
                )
                .mean()
            )

            pokupk_rolling_std = (
                group["POKUPK_DAILY_TURNOVER"]
                .rolling(
                    window=rolling_window,
                    min_periods=2,
                )
                .std()
            )

            prodano_rolling_mean = (
                group["PRODANO_DAILY_TURNOVER"]
                .rolling(
                    window=rolling_window,
                    min_periods=2,
                )
                .mean()
            )

            prodano_rolling_std = (
                group["PRODANO_DAILY_TURNOVER"]
                .rolling(
                    window=rolling_window,
                    min_periods=2,
                )
                .std()
            )

            group["POKUPK_7D_VOL_PCT"] = (
                    pokupk_rolling_std / pokupk_rolling_mean.abs() * 100
            )

            group["PRODANO_7D_VOL_PCT"] = (
                    prodano_rolling_std / prodano_rolling_mean.abs() * 100
            )

            group["POKUPK_7D_VOL_PCT"] = (
                group["POKUPK_7D_VOL_PCT"]
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0)
            )

            group["PRODANO_7D_VOL_PCT"] = (
                group["PRODANO_7D_VOL_PCT"]
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0)
            )

            return group

        indicator_parts = []

        for _, group in indicators.groupby(
                [
                    "BANK_NAME",
                    "VAL_NAME",
                ],
                sort=False,
        ):
            indicator_parts.append(
                add_rolling_volatility(group)
            )

        indicators = pd.concat(
            indicator_parts,
            ignore_index=True,
        )

        return indicators

    def calculate_additional_currency_indicators(df_indicators, rolling_window=7):
            # -------------------------------------------------
            # Daily average transaction size and transaction count
            # by currency
            # -------------------------------------------------
            df_currency_daily = (
                df_indicators
                .groupby(["DATVAL", "VAL_NAME"], as_index=False)
                .agg(
                    TOTAL_DAILY_TURNOVER=("TOTAL_DAILY_TURNOVER", "sum"),
                    TRANSACTION_COUNT=("TRANSACTION_COUNT", "sum"),
                )
            )

            df_currency_daily["AVG_TRANSACTION_SIZE"] = np.where(
                df_currency_daily["TRANSACTION_COUNT"] > 0,
                df_currency_daily["TOTAL_DAILY_TURNOVER"]
                / df_currency_daily["TRANSACTION_COUNT"],
                0,
            )

            df_currency_daily = df_currency_daily.sort_values(
                [
                    "VAL_NAME",
                    "DATVAL",
                ]
            )

            df_currency_daily["AVG_TRANSACTION_SIZE_7D"] = (
                df_currency_daily
                .groupby("VAL_NAME")["AVG_TRANSACTION_SIZE"]
                .transform(
                    lambda x: x.rolling(
                        window=rolling_window,
                        min_periods=1,
                    ).mean()
                )
            )

            # -------------------------------------------------
            # Concentration indicator, HHI
            # Calculated across banks for each currency and day.
            # -------------------------------------------------
            df_bank_currency_daily = (
                df_indicators
                .groupby(["DATVAL", "VAL_NAME", "BANK_NAME"], as_index=False)
                .agg(
                    BANK_DAILY_TURNOVER=("TOTAL_DAILY_TURNOVER", "sum"),
                )
            )

            df_bank_currency_daily["CURRENCY_DAILY_TURNOVER"] = (
                df_bank_currency_daily
                .groupby(["DATVAL", "VAL_NAME"])["BANK_DAILY_TURNOVER"]
                .transform("sum")
            )

            df_bank_currency_daily["BANK_SHARE"] = np.where(
                df_bank_currency_daily["CURRENCY_DAILY_TURNOVER"] > 0,
                df_bank_currency_daily["BANK_DAILY_TURNOVER"]
                / df_bank_currency_daily["CURRENCY_DAILY_TURNOVER"],
                0,
            )

            df_bank_currency_daily["BANK_SHARE_SQUARED"] = (
                    df_bank_currency_daily["BANK_SHARE"] ** 2
            )

            df_concentration = (
                df_bank_currency_daily
                .groupby(["DATVAL", "VAL_NAME"], as_index=False)
                .agg(
                    CONCENTRATION_HHI=("BANK_SHARE_SQUARED", "sum"),
                )
            )

            df_concentration["CONCENTRATION_HHI"] = (
                    df_concentration["CONCENTRATION_HHI"] * 10000
            )

            df_additional_indicators = df_currency_daily.merge(
                df_concentration,
                on=["DATVAL", "VAL_NAME"],
                how="left",
            )

            df_additional_indicators["CONCENTRATION_HHI"] = (
                df_additional_indicators["CONCENTRATION_HHI"]
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0)
            )

            return df_additional_indicators


    def weighted_average_indicator_rate(group, rate_column, weight_column):
        rates = pd.to_numeric(
            group[rate_column],
            errors="coerce",
        )

        weights = pd.to_numeric(
            group[weight_column],
            errors="coerce",
        ).fillna(0).abs()

        valid_mask = (
                rates.notna()
                & weights.notna()
                & (weights > 0)
                & (rates > 0)
        )

        if valid_mask.any():
            return np.average(
                rates.loc[valid_mask],
                weights=weights.loc[valid_mask],
            )

        valid_rates = rates[
            rates.notna()
            & (rates > 0)
            ]

        if not valid_rates.empty:
            return valid_rates.mean()

        return np.nan


    def calculate_bid_ask_spread_indicators(df_source):
        group_columns = [
            "DATVAL_DAY",
            "VAL_NAME",
        ]

        spread_rows = []

        for group_values, group_df in df_source.groupby(group_columns):
            datval_day, val_name = group_values

            kur_pok_weighted = weighted_average_indicator_rate(
                group_df,
                "KUR_POK",
                "POKUPK",
            )

            kur_pr_weighted = weighted_average_indicator_rate(
                group_df,
                "KUR_PR",
                "PRODANO",
            )

            bid_ask_spread = kur_pr_weighted - kur_pok_weighted

            bid_ask_spread_pct = (
                bid_ask_spread / kur_pok_weighted * 100
                if pd.notna(kur_pok_weighted) and kur_pok_weighted > 0
                else np.nan
            )

            spread_rows.append(
                {
                    "DATVAL": datval_day,
                    "VAL_NAME": val_name,
                    "KUR_POK_WEIGHTED": kur_pok_weighted,
                    "KUR_PR_WEIGHTED": kur_pr_weighted,
                    "BID_ASK_SPREAD": bid_ask_spread,
                    "BID_ASK_SPREAD_PCT": bid_ask_spread_pct,
                }
            )

        df_spread = pd.DataFrame(spread_rows)

        if df_spread.empty:
            return df_spread

        df_spread["BID_ASK_SPREAD"] = (
            df_spread["BID_ASK_SPREAD"]
            .replace([np.inf, -np.inf], np.nan)
        )

        df_spread["BID_ASK_SPREAD_PCT"] = (
            df_spread["BID_ASK_SPREAD_PCT"]
            .replace([np.inf, -np.inf], np.nan)
        )

        return df_spread.sort_values(
            [
                "VAL_NAME",
                "DATVAL",
            ]
        )

    def format_number(value):
        return f"{value:,.2f}"



    def render_indicator_section(
        df_source,
        section_title,
        chart_dimension,
    ):
        if df_source.empty:
            st.warning(f"Нет данных для раздела: {section_title}")
            return

        df_indicators = calculate_daily_fx_indicators(
            df_source,
            rolling_window=ROLLING_VOL_WINDOW,
        )

        if df_indicators.empty:
            st.warning(f"Не удалось рассчитать индикаторы для раздела: {section_title}")
            return

        st.subheader(section_title)

        # -------------------------------------------------
        # KPI cards
        # Turnover KPIs are calculated as daily averages
        # for the selected indication period.
        # -------------------------------------------------
        df_daily_kpi = (
            df_indicators
            .groupby("DATVAL", as_index=False)
            .agg(
                POKUPK_DAILY_TOTAL=("POKUPK_DAILY_TURNOVER", "sum"),
                PRODANO_DAILY_TOTAL=("PRODANO_DAILY_TURNOVER", "sum"),
                TOTAL_DAILY_TOTAL=("TOTAL_DAILY_TURNOVER", "sum"),
            )
        )

        avg_daily_pokupk = df_daily_kpi["POKUPK_DAILY_TOTAL"].mean()
        avg_daily_prodano = df_daily_kpi["PRODANO_DAILY_TOTAL"].mean()
        avg_daily_turnover = df_daily_kpi["TOTAL_DAILY_TOTAL"].mean()

        avg_pokupk_vol = df_indicators["POKUPK_7D_VOL_PCT"].mean()
        avg_prodano_vol = df_indicators["PRODANO_7D_VOL_PCT"].mean()

        kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)

        with kpi_col1:
            st.metric(
                "Покупка - среднедневной оборот",
                format_number(avg_daily_pokupk),
            )

        with kpi_col2:
            st.metric(
                "Продажа - среднедневной оборот",
                format_number(avg_daily_prodano),
            )

        with kpi_col3:
            st.metric(
                "Среднедневной оборот - Всего",
                format_number(avg_daily_turnover),
            )

        with kpi_col4:
            st.metric(
                "Покупка - 7-дневная волатильность",
                f"{avg_pokupk_vol:,.2f}%",
            )

        with kpi_col5:
            st.metric(
                "Продажа - 7-дневная волатильность",
                f"{avg_prodano_vol:,.2f}%",
            )

        # -------------------------------------------------
        # Prepare chart data
        # -------------------------------------------------
        chart_color_col = (
            "VAL_NAME"
            if chart_dimension == "Валюта"
            else "BANK_NAME"
        )

        df_turnover_chart = (
            df_indicators
            .groupby(["DATVAL", chart_color_col], as_index=False)
            .agg(
                POKUPK_DAILY_TURNOVER=("POKUPK_DAILY_TURNOVER", "sum"),
                PRODANO_DAILY_TURNOVER=("PRODANO_DAILY_TURNOVER", "sum"),
                TOTAL_DAILY_TURNOVER=("TOTAL_DAILY_TURNOVER", "sum"),
            )
        )

        df_volatility_chart = (
            df_indicators
            .groupby(["DATVAL", chart_color_col], as_index=False)
            .agg(
                POKUPK_7D_VOL_PCT=("POKUPK_7D_VOL_PCT", "mean"),
                PRODANO_7D_VOL_PCT=("PRODANO_7D_VOL_PCT", "mean"),
            )
        )

        df_swap_chart = (
            df_indicators
            .groupby(["DATVAL", chart_color_col], as_index=False)
            .agg(
                POS_SVOP1_DAILY_TURNOVER=("POS_SVOP1_DAILY_TURNOVER", "sum"),
                POS_SVOP2_DAILY_TURNOVER=("POS_SVOP2_DAILY_TURNOVER", "sum"),
            )
        )

        df_additional_indicators = calculate_additional_currency_indicators(
            df_indicators,
            rolling_window=ROLLING_VOL_WINDOW,
        )

        df_bid_ask_spread = calculate_bid_ask_spread_indicators(
            df_source
        )

        show_swap_tab = section_title == "Безналичные операции"

        if show_swap_tab:
            turnover_tab, volatility_tab, additional_tab, swap_tab = st.tabs(
                [
                    "Ежедневный оборот",
                    "7-дневная волатильность",
                    "Дополнительные индикаторы",
                    "СВОП-операции",
                ]
            )
        else:
            turnover_tab, volatility_tab, additional_tab = st.tabs(
                [
                    "Ежедневный оборот",
                    "7-дневная волатильность",
                    "Дополнительные индикаторы",
                ]
            )

        # -------------------------------------------------
        # Daily turnover charts
        # -------------------------------------------------
        with turnover_tab:
            fig_total_turnover = px.bar(
                df_turnover_chart,
                x="DATVAL",
                y="TOTAL_DAILY_TURNOVER",
                color=chart_color_col,
                orientation="v",
                title=f"<b>{section_title}: Общий ежедневный оборот — покупка + продажа</b>",
                template="plotly_white",
            )

            fig_total_turnover.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(showgrid=False),
                yaxis_title="Ежедневный оборот",
            )

            st.plotly_chart(
                fig_total_turnover,
                use_container_width=True,
            )

            turnover_col1, turnover_col2 = st.columns(2)

            with turnover_col1:
                fig_pokupk_turnover = px.bar(
                    df_turnover_chart,
                    x="DATVAL",
                    y="POKUPK_DAILY_TURNOVER",
                    color=chart_color_col,
                    orientation="v",
                    title=f"<b>{section_title}: Ежедневный оборот — покупка</b>",
                    template="plotly_white",
                )

                fig_pokupk_turnover.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="Покупка",
                )

                st.plotly_chart(
                    fig_pokupk_turnover,
                    use_container_width=True,
                )

            with turnover_col2:
                fig_prodano_turnover = px.bar(
                    df_turnover_chart,
                    x="DATVAL",
                    y="PRODANO_DAILY_TURNOVER",
                    color=chart_color_col,
                    orientation="v",
                    title=f"<b>{section_title}: Ежедневный оборот — продажа</b>",
                    template="plotly_white",
                )

                fig_prodano_turnover.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="Продажа",
                )

                st.plotly_chart(
                    fig_prodano_turnover,
                    use_container_width=True,
                )

        # -------------------------------------------------
        # 7-day volatility charts
        # -------------------------------------------------
        with volatility_tab:
            volatility_col1, volatility_col2 = st.columns(2)

            with volatility_col1:
                fig_pokupk_volatility = px.line(
                    df_volatility_chart,
                    x="DATVAL",
                    y="POKUPK_7D_VOL_PCT",
                    color=chart_color_col,
                    markers=True,
                    title=f"<b>{section_title}: 7-дневная волатильность — покупка</b>",
                    template="plotly_white",
                )

                fig_pokupk_volatility.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="7-дневная волатильность, %",
                )

                st.plotly_chart(
                    fig_pokupk_volatility,
                    use_container_width=True,
                )

            with volatility_col2:
                fig_prodano_volatility = px.line(
                    df_volatility_chart,
                    x="DATVAL",
                    y="PRODANO_7D_VOL_PCT",
                    color=chart_color_col,
                    markers=True,
                    title=f"<b>{section_title}: 7-дневная волатильность — продажа</b>",
                    template="plotly_white",
                )

                fig_prodano_volatility.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="7-дневная волатильность, %",
                )

                st.plotly_chart(
                    fig_prodano_volatility,
                    use_container_width=True,
                )

        # -------------------------------------------------
        # Additional FX market indicators
        # -------------------------------------------------
        with additional_tab:
            additional_col1, additional_col2 = st.columns(2)

            with additional_col1:
                fig_avg_transaction_size = px.line(
                    df_additional_indicators,
                    x="DATVAL",
                    y="AVG_TRANSACTION_SIZE_7D",
                    color="VAL_NAME",
                    markers=True,
                    title="<b>7-дневный средний размер валютной операции</b>",
                    template="plotly_white",
                )

                fig_avg_transaction_size.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="Средний размер операции",
                    legend_title_text="Валюта",
                )

                st.plotly_chart(
                    fig_avg_transaction_size,
                    use_container_width=True,
                )

            with additional_col2:
                fig_transaction_count = px.bar(
                    df_additional_indicators,
                    x="DATVAL",
                    y="TRANSACTION_COUNT",
                    color="VAL_NAME",
                    orientation="v",
                    title="<b>Количество валютных операций</b>",
                    template="plotly_white",
                )

                fig_transaction_count.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="Количество операций",
                    legend_title_text="Валюта",
                )

                st.plotly_chart(
                    fig_transaction_count,
                    use_container_width=True,
                )

            fig_concentration = px.line(
                df_additional_indicators,
                x="DATVAL",
                y="CONCENTRATION_HHI",
                color="VAL_NAME",
                markers=True,
                title="<b>Индикатор концентрации валютного рынка, HHI</b>",
                template="plotly_white",
            )

            fig_concentration.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(showgrid=False),
                yaxis_title="HHI",
                legend_title_text="Валюта",
            )

            st.plotly_chart(
                fig_concentration,
                use_container_width=True,
            )

            st.markdown(
                """
                <div style="
                    background-color: rgba(240, 242, 246, 0.08);
                    border-left: 4px solid #4C78A8;
                    padding: 14px 18px;
                    border-radius: 6px;
                    margin-top: 8px;
                    margin-bottom: 18px;
                    font-size: 0.95rem;
                    line-height: 1.6;
                ">
                    <b>Индекс Херфиндаля-Хиршмана (HHI)</b> – это показатель,
                    измеряющий уровень концентрации рынка и степень конкуренции между
                    компаниями в отрасли. Он рассчитывается как сумма квадратов рыночных
                    долей всех участников рынка.
                    <br><br>
                    Значение индекса варьируется от <b>0</b> — абсолютная конкуренция —
                    до <b>10000</b> — абсолютная монополия.
                    <br><br>
                    <b>Стандартная шкала интерпретации HHI:</b>
                    <ul>
                        <li><b>менее 1500:</b> рынок с низкой концентрацией
                        — высококонкурентный. В этой среде ни одна компания не имеет
                        доминирующего влияния.</li>
                        <li><b>от 1500 до 2500:</b> рынок с умеренной концентрацией.
                        Нормальная рыночная среда с достаточным количеством игроков
                        и средним уровнем конкуренции.</li>
                        <li><b>более 2500:</b> рынок с высокой концентрацией
                        — олигополия или монополия. Присутствует явный лидер или несколько
                        крупных игроков, что увеличивает риск нарушения антимонопольного
                        законодательства.</li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown("---")

            fig_bid_ask_spread = px.line(
                df_bid_ask_spread.dropna(subset=["BID_ASK_SPREAD"]),
                x="DATVAL",
                y="BID_ASK_SPREAD",
                color="VAL_NAME",
                markers=True,
                title="<b>Bid-ask spread - курс продажи минус курс покупки</b>",
                template="plotly_white",
            )

            fig_bid_ask_spread.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(showgrid=False),
                yaxis_title="Bid-ask spread",
                legend_title_text="Валюта",
            )

            st.plotly_chart(
                fig_bid_ask_spread,
                use_container_width=True,
            )

            fig_bid_ask_spread_pct = px.line(
                df_bid_ask_spread.dropna(subset=["BID_ASK_SPREAD_PCT"]),
                x="DATVAL",
                y="BID_ASK_SPREAD_PCT",
                color="VAL_NAME",
                markers=True,
                title="<b>Bid-ask spread - % от курса покупки</b>",
                template="plotly_white",
            )

            fig_bid_ask_spread_pct.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(showgrid=False),
                yaxis_title="Bid-ask spread, %",
                legend_title_text="Валюта",
            )

            st.plotly_chart(
                fig_bid_ask_spread_pct,
                use_container_width=True,
            )

        # -------------------------------------------------
        # Swap operation charts
        # Only for non-cash operations
        # -------------------------------------------------
        if show_swap_tab:
            with swap_tab:
                fig_svop_inflow = px.bar(
                    df_swap_chart,
                    x="DATVAL",
                    y="POS_SVOP1_DAILY_TURNOVER",
                    color=chart_color_col,
                    orientation="v",
                    title="<b>Поступление валюты от СВОП-операций</b>",
                    template="plotly_white",
                )

                fig_svop_inflow.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="Поступление валюты",
                )

                st.plotly_chart(
                    fig_svop_inflow,
                    use_container_width=True,
                )

                fig_svop_outflow = px.bar(
                    df_swap_chart,
                    x="DATVAL",
                    y="POS_SVOP2_DAILY_TURNOVER",
                    color=chart_color_col,
                    orientation="v",
                    title="<b>Отток валюты по СВОП-операциям</b>",
                    template="plotly_white",
                )

                fig_svop_outflow.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=False),
                    yaxis_title="Отток валюты",
                )

                st.plotly_chart(
                    fig_svop_outflow,
                    use_container_width=True,
                )


    # -------------------------------------------------
    # Build separate NonCash and Cash datasets
    # -------------------------------------------------
    df_noncash_indicators = build_indicator_base(
        df_noncash,
        "Безналичные операции",
    )

    df_cash_indicators = build_indicator_base(
        df_cash,
        "Наличные операции",
    )

    df_fx_base = pd.concat(
        [
            df_noncash_indicators,
            df_cash_indicators,
        ],
        ignore_index=True,
    )

    if df_fx_base.empty:
        st.warning("Нет данных для расчета индикаторов валютного рынка.")
        st.stop()

    # -------------------------------------------------
    # Sidebar filters
    # -------------------------------------------------
    min_date = df_fx_base["DATVAL"].min().date()
    max_date = df_fx_base["DATVAL"].max().date()

    start_date = st.sidebar.date_input(
        "Выберите начальную дату:",
        min_value=min_date,
        max_value=max_date,
        value=min_date,
        key="fx_indicators_start_date",
    )

    end_date = st.sidebar.date_input(
        "Выберите конечную дату:",
        min_value=min_date,
        max_value=max_date,
        value=max_date,
        key="fx_indicators_end_date",
    )

    if start_date > end_date:
        st.sidebar.error("Начальная дата не может быть позже конечной даты.")
        st.stop()

    indicator_bank_options = sorted(
        df_fx_base["BANK_NAME"].dropna().unique()
    )

    indicator_currency_options = sorted(
        df_fx_base["VAL_NAME"].dropna().unique()
    )

    indicator_currency_options = move_option_to_front(
        indicator_currency_options,
        "USD",
    )

    default_currencies = (
        ["USD"]
        if "USD" in indicator_currency_options
        else indicator_currency_options[:1]
    )

    selected_banks = st.sidebar.multiselect(
        "Выберите Банк:",
        options=indicator_bank_options,
        default=indicator_bank_options,
        key="fx_indicators_banks",
    )

    selected_currencies = st.sidebar.multiselect(
        "Выберите Валюту:",
        options=indicator_currency_options,
        default=default_currencies,
        key="fx_indicators_currencies",
    )

    chart_dimension = st.sidebar.radio(
        "Разрез для графиков:",
        options=[
            "Валюта",
            "Банк",
        ],
        index=0,
        key="fx_indicators_chart_dimension",
    )

    if not selected_banks:
        st.warning("Выберите хотя бы один банк.")
        st.stop()

    if not selected_currencies:
        st.warning("Выберите хотя бы одну валюту.")
        st.stop()

    # -------------------------------------------------
    # Apply common filters separately
    # -------------------------------------------------
    start_date = pd.Timestamp(start_date)
    end_date = pd.Timestamp(end_date)

    def apply_indicator_filters(df):
        return df[
            df["DATVAL"].between(start_date, end_date)
            & df["BANK_NAME"].isin(selected_banks)
            & df["VAL_NAME"].isin(selected_currencies)
        ].copy()


    df_noncash_filtered = apply_indicator_filters(
        df_noncash_indicators
    )

    df_cash_filtered = apply_indicator_filters(
        df_cash_indicators
    )

    # -------------------------------------------------
    # Render separate tabs for NonCash and Cash
    # -------------------------------------------------
    noncash_tab, cash_tab = st.tabs(
        [
            "Безналичные операции",
            "Наличные операции",
        ]
    )

    with noncash_tab:
        render_indicator_section(
            df_source=df_noncash_filtered,
            section_title="Безналичные операции",
            chart_dimension=chart_dimension,
        )

    with cash_tab:
        render_indicator_section(
            df_source=df_cash_filtered,
            section_title="Наличные операции",
            chart_dimension=chart_dimension,
        )