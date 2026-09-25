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

    # --- Отрезаем футер (всё после слова "Спасибо,") ---
    cutoff = len(words)
    for i, w in enumerate(words):
        if w["text"].strip() == "Спасибо,":
            cutoff = i
            break
    words = words[:cutoff]

    # --- ШАГ 1. Находим координаты колонок из шапки ---
    x_income = find_column_x(words, "Приход")
    x_expense = find_column_x(words, "Расход")

    if x_income is None or x_expense is None:
        return []

    # --- ШАГ 2. Находим даты-маркеры (x < 40) ---
    date_re = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    markers = []
    for i, w in enumerate(words):
        if date_re.match(w["text"]) and w["x0"] < 40:
            markers.append(i)

    if not markers:
        return []

    # --- ШАГ 3. Режем на блоки между маркерами ---
    rows = []
    for k, start_idx in enumerate(markers):
        end_idx = markers[k + 1] if k + 1 < len(markers) else len(words)
        block = words[start_idx:end_idx]
        row = parse_block(block, x_income, x_expense)
        if row:
            rows.append(row)
    return rows


def parse_block(block, x_income, x_expense):
    """block — список слов одной операции."""
    date = None
    income = None
    expense = None
    desc_words = []

    # Границы колонок (расширенные, чтобы ловить сдвиги)
    income_lo, income_hi = x_income - 25, x_income + 20
    expense_lo, expense_hi = x_expense - 20, x_expense + 25

    # Слова-футер на всякий случай
    FOOTER_WORDS = {
        "Спасибо,", "что", "Вы", "с", "нами!", "Всегда", "Ваш,",
        "Банк", "ВТБ", "(ПАО)", "RUB", "₽",
    }

    for w in block:
        text = w["text"]
        x = w["x0"]

        # Пропускаем футерные слова
        if text in FOOTER_WORDS:
            continue

        # Дата (первое слово x<40)
        if date is None and re.match(r"^\d{2}\.\d{2}\.\d{4}$", text) and x < 40:
            date = text
            continue

        # Пропускаем время (x<40)
        if re.match(r"^\d{2}:\d{2}:\d{2}$", text) and x < 40:
            continue

        # Суммы: приход/расход по координате колонки
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

        # Описание — всё, что правее колонки "Расход" + 20
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
