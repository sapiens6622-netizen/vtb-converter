import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# ============ НАСТРОЙКИ ============
COLUMNS = ["Дата", "Приход", "Расход", "Комментарий"]


# ============ ИЗВЛЕЧЕНИЕ БАЛАНСОВ ============
def extract_metadata(pdf) -> dict:
    """Вытаскиваем только два баланса с 1-й страницы."""
    meta = {}
    text = pdf.pages[0].extract_text() or ""

    for label in [
        "Баланс на начало периода",
        "Баланс на конец периода",
    ]:
        m = re.search(re.escape(label) + r"\s+([\d\s,]+\.\d{2})\s*RUB", text)
        if m:
            meta[label] = parse_money(m.group(1) + " RUB")

    return meta


# ============ ФУНКЦИИ ПАРСИНГА ============
def clean(s):
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def parse_money(s: str):
    """'4,000.00' -> 4000.0; '0' -> 0.0; '' -> None"""
    if not s:
        return None
    s = s.replace("RUB", "").replace("₽", "").replace(",", "").replace(" ", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def find_column_x(words, column_name: str):
    """Ищем слово column_name в шапке и возвращаем его x0."""
    for w in words:
        if w["text"].strip().lower() == column_name.lower():
            return w["x0"]
    return None


def parse_operations(page) -> list:
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    if not words:
        return []

    # --- Отрезаем футер ---
    cutoff = len(words)
    for i, w in enumerate(words):
        if w["text"].strip() == "Спасибо,":
            cutoff = i
            break
    words = words[:cutoff]

    # --- Находим координаты колонок ---
    x_income = find_column_x(words, "Приход")
    x_expense = find_column_x(words, "Расход")

    if x_income is None or x_expense is None:
        return []

    # --- Находим все даты-маркеры ---
    date_re = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    date_markers = []
    for i, w in enumerate(words):
        if date_re.match(w["text"]) and w["x0"] < 40:
            date_markers.append((i, w["top"]))

    if not date_markers:
        return []

    # --- Уникализация по top ---
    boundaries = []
    seen_tops = set()
    for idx, top in date_markers:
        top_rounded = round(top)
        if top_rounded in seen_tops:
            continue
        seen_tops.add(top_rounded)
        boundaries.append((idx, top_rounded))

    if not boundaries:
        return []

    # --- Парсим каждый блок ---
    rows = []
    for k, (start_idx, start_top) in enumerate(boundaries):
        if k + 1 < len(boundaries):
            end_idx = boundaries[k + 1][0]
            end_top = boundaries[k + 1][1]
        else:
            end_idx = len(words)
            end_top = 99999

        block = words[start_idx:end_idx]
        row = parse_block(block, x_income, x_expense, end_top)
        if row:
            rows.append(row)
    return rows


def parse_block(block, x_income, x_expense, end_top):
    """block — список слов одной операции."""
    date = None
    income = None
    expense = None
    desc_words = []

    # Расширенные границы колонок (расход расширен влево)
    income_lo, income_hi = x_income - 30, x_income + 25
    expense_lo, expense_hi = x_expense - 35, x_expense + 30

    FOOTER_WORDS = {
        "Спасибо,", "что", "Вы", "с", "нами!", "Всегда", "Ваш,",
        "Банк", "ВТБ", "(ПАО)", "₽",
    }
    HEADER_WORDS = {
        "Дата", "и", "время", "обработки", "банком", "Сумма",
        "операции", "в", "валюте", "счета/карты", "Комиссия",
        "Описание", "Приход", "Расход", "Наименование",
        "получателя/", "Отправителя",
    }

    for w in block:
        text = w["text"]
        x = w["x0"]

        if text in FOOTER_WORDS:
            continue
        if text in HEADER_WORDS and date is None:
            continue
        if text == "RUB":
            continue

        # Дата
        if date is None and re.match(r"^\d{2}\.\d{2}\.\d{4}$", text) and x < 40:
            date = text
            continue

        if re.match(r"^\d{2}:\d{2}:\d{2}$", text) and x < 40:
            continue

        if re.match(r"^\d{2}\.\d{2}\.\d{4}$", text) and 50 < x < 130:
            continue

        # Суммы
        if re.match(r"^-?[\d,]+\.\d{2}$", text):
            value = parse_money(text)
            if value is not None:
                if income_lo <= x <= income_hi:
                    if value != 0:
                        income = abs(value)
                    continue
                if expense_lo <= x <= expense_hi:
                    if value != 0:
                        expense = abs(value)
                    continue
            continue

        # Описание
        if x > x_expense + 20:
            if re.match(r"^\d{1,2}$", text):
                continue
            desc_words.append(text)

    if not date:
        return None

    desc = " ".join(desc_words)
    desc = re.sub(r"\s+", " ", desc).strip()

    return [date, income, expense, desc]


# ============ ИНТЕРФЕЙС STREAMLIT ============
st.set_page_config(page_title="Конвертер выписок ВТБ", page_icon="📄")
st.title("📄 Конвертер выписок ВТБ в Excel")
st.markdown("Загрузите PDF-выписку ВТБ, чтобы получить таблицу операций в формате Excel.")

uploaded_file = st.file_uploader("Выберите PDF-файл выписки", type=["pdf"])

if uploaded_file is not None:
    if st.button("Конвертировать в Excel"):
        with st.spinner("Обрабатываю документ..."):
            try:
                with open("temp_uploaded.pdf", "wb") as f:
                    f.write(uploaded_file.getbuffer())

                all_rows = []
                meta = {}
                with pdfplumber.open("temp_uploaded.pdf") as pdf:
                    meta = extract_metadata(pdf)
                    for page_num, page in enumerate(pdf.pages, start=1):
                        rows = parse_operations(page)
                        all_rows.extend(rows)

                if not all_rows:
                    st.warning("Не удалось найти операции. Проверьте формат PDF.")
                else:
                    df = pd.DataFrame(all_rows, columns=COLUMNS)

                    st.subheader("📋 Информация о счёте")

                    info_col1, info_col2 = st.columns(2)
                    with info_col1:
                        if "Баланс на начало периода" in meta:
                            st.markdown(
                                f"**Баланс на начало периода:** "
                                f"{meta['Баланс на начало периода']:.2f} RUB"
                            )
                    with info_col2:
                        if "Баланс на конец периода" in meta:
                            st.markdown(
                                f"**Баланс на конец периода:** "
                                f"{meta['Баланс на конец периода']:.2f} RUB"
                            )

                    st.markdown("---")
                    st.subheader("📊 Операции по счёту")

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        if meta:
                            meta_rows = []
                            for label in [
                                "Баланс на начало периода",
                                "Баланс на конец периода",
                            ]:
                                if label in meta:
                                    meta_rows.append([label, meta[label]])

                            meta_df = pd.DataFrame(
                                meta_rows, columns=["Параметр", "Значение"]
                            )
                            meta_df.to_excel(
                                writer, index=False, sheet_name="Информация о счёте"
                            )
                            ws_meta = writer.sheets["Информация о счёте"]
                            ws_meta.column_dimensions["A"].width = 30
                            ws_meta.column_dimensions["B"].width = 20

                        df.to_excel(writer, index=False, sheet_name="Операции")
                        ws = writer.sheets["Операции"]

                        for col, w in zip("ABCD", [14, 14, 14, 70]):
                            ws.column_dimensions[col].width = w

                        for cell in ws["A"][1:]:
                            cell.number_format = "@"

                        for col in ("B", "C"):
                            for cell in ws[col][1:]:
                                cell.number_format = "0"

                    processed_data = output.getvalue()

                    st.success(f"Готово! Найдено операций: {len(df)}")
                    st.dataframe(df)

                    st.download_button(
                        label="⬇️ Скачать Excel",
                        data=processed_data,
                        file_name="выписка_втб.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                st.error(f"Ошибка при обработке: {e}")
