import streamlit as st
import pandas as pd
from datetime import datetime
import tempfile
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="АПП Генератор ПВЗ", layout="wide")
st.title("🛠 Генератор АПП для ПВЗ")

# Инициализация
if 'scanned' not in st.session_state:
    st.session_state.scanned = []
if 'df' not in st.session_state:
    st.session_state.df = None
if 'fio' not in st.session_state:
    st.session_state.fio = ""

# ====================== БОКОВАЯ ПАНЕЛЬ ======================
with st.sidebar:
    st.header("Смена")
    fio = st.text_input("Ваше ФИО", value=st.session_state.fio, placeholder="Фамилия Имя Отчество")
    if fio:
        st.session_state.fio = fio

    st.divider()
    uploaded_file = st.file_uploader("1. Загрузить CSV", type=['csv'])
    
    if uploaded_file:
        try:
            sample = uploaded_file.read(2048).decode('utf-8', errors='ignore')
            uploaded_file.seek(0)
            
            if 'sep=' in sample.lower():
                df = pd.read_csv(uploaded_file, sep=None, engine='python', encoding='utf-8', skiprows=1)
            else:
                try:
                    df = pd.read_csv(uploaded_file, sep=',', encoding='utf-8')
                except:
                    df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            
            df.columns = [str(col).strip().replace('"', '').replace("'", "") for col in df.columns]
            st.session_state.df = df
            st.success(f"✅ Загружено {len(df)} строк")
            
        except Exception as e:
            st.error(f"❌ Ошибка чтения файла: {e}")

# ====================== СКАНИРОВАНИЕ ======================
col1, col2 = st.columns([3, 1])

with col1:
    st.subheader("2. Сканирование заказов")
    with st.form(key="scan_form", clear_on_submit=True):
        order_input = st.text_input("Номер заказа", placeholder="260597447132 или LM-260597447132")
        submitted = st.form_submit_button("✅ Добавить", type="primary", use_container_width=True)
        
        if submitted and order_input.strip():
            if st.session_state.df is None:
                st.error("Сначала загрузите CSV файл!")
            else:
                raw_num = str(order_input).strip()
                search_num = raw_num.replace("LM-", "").replace("lm-", "").strip()
                
                df = st.session_state.df
                found = False
                order_columns = [col for col in df.columns if any(word in col.lower() 
                                for word in ['заказ', 'order', 'номер'])]
                if not order_columns:
                    order_columns = df.columns

                for col in order_columns:
                    df[col] = df[col].astype(str).str.strip()
                    result = df[df[col] == search_num]
                    if not result.empty:
                        row = result.iloc[0].to_dict()
                        if row not in st.session_state.scanned:
                            st.session_state.scanned.append(row)
                            st.success(f"✅ УСПЕШНО ДОБАВЛЕН: {search_num}")
                        else:
                            st.error("🔴 ДУБЛИКАТ!")
                        found = True
                        break
                if not found:
                    st.error(f"❌ Заказ {search_num} не найден")

with col2:
    st.subheader("Статистика")
    st.metric("Отсканировано", len(st.session_state.scanned))

# ====================== НЕОТСКАНИРОВАННЫЕ ======================
if st.session_state.df is not None:
    st.subheader("📋 Неотсканированные заказы")
    df = st.session_state.df.copy()
    if "Статус заказа" in df.columns and "Статус контроля" in df.columns:
        mask = (
            df["Статус заказа"].astype(str).str.contains('Собран', na=False) &
            df["Статус контроля"].astype(str).str.contains('пройден', na=False)
        )
        remaining = df[mask].copy()
        
        if st.session_state.scanned:
            scanned_set = {str(r.get('Заказ', '')) for r in st.session_state.scanned}
            remaining = remaining[~remaining['Заказ'].astype(str).isin(scanned_set)]
        
        st.write(f"**Осталось отсканировать: {len(remaining)}**")
        if not remaining.empty:
            st.dataframe(remaining, use_container_width=True, hide_index=True)

# ====================== ОТСКАНИРОВАННЫЕ ======================
st.subheader(f"✅ Отсканированные ({len(st.session_state.scanned)})")
if st.session_state.scanned:
    st.dataframe(pd.DataFrame(st.session_state.scanned), use_container_width=True, hide_index=True)

# ====================== ФОРМИРОВАНИЕ ОДНОГО ФАЙЛА ======================
if st.button("🚀 Сформировать все АПП", type="secondary", use_container_width=True):
    if not st.session_state.scanned:
        st.error("Нет отсканированных заказов")
    else:
        scanned_df = pd.DataFrame(st.session_state.scanned)
        current_date = datetime.now().strftime("%d.%m.%Y")
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            path = tmp.name
        
        try:
            wb = Workbook()
            wb.remove(wb.active)
            
            partner_names = {
                '135_FIVEPOST': '5Post',
                '135_SDEK': 'SDEK',
                '135_Yandex': 'Yandex',
                'UNKNOWN': 'DPD'
            }
            
            for code, sheet_name in partner_names.items():
                group = scanned_df[scanned_df.get('Партнер', '') == code]
                if group.empty:
                    continue
                
                ws = wb.create_sheet(title=sheet_name)
                
                ws['A1'] = "Акт приема-передачи"
                ws.merge_cells('A1:D1')
                ws['A1'].font = Font(bold=True, size=14)
                ws['A1'].alignment = Alignment(horizontal="center")
                
                ws['A2'] = f"от {current_date}"
                ws.merge_cells('A2:D2')
                ws['A2'].alignment = Alignment(horizontal="center")
                
                # Таблица
                table_start = 5
                for col, header in enumerate(["№", "номер отправления", "вес отправления (кг)", "стоимость отправления(руб,)"], 1):
                    cell = ws.cell(row=table_start, column=col, value=header)
                    cell.font = Font(bold=True)
                
                total_w = 0
                total_c = 0
                for i, row in enumerate(group.itertuples(index=False), 1):
                    row_dict = dict(zip(group.columns, row))
                    r = table_start + i
                    ws.cell(row=r, column=1, value=i)
                    ws.cell(row=r, column=2, value=str(row_dict.get('Заказ', '')))
                    w = float(row_dict.get('Вес', 0))
                    c = float(row_dict.get('Цена заказа', 0))
                    ws.cell(row=r, column=3, value=round(w, 3))
                    ws.cell(row=r, column=4, value=int(c))
                    total_w += w
                    total_c += c
                
                # Итого
                last = table_start + len(group)
                ws.cell(row=last+1, column=1, value="Итого:")
                ws.cell(row=last+1, column=2, value=len(group))
                ws.cell(row=last+1, column=3, value=round(total_w, 3))
                ws.cell(row=last+1, column=4, value=int(total_c))
            
            wb.save(path)
            
            with open(path, 'rb') as f:
                st.download_button(
                    label="📥 Скачать АПП (все партнёры в одном файле)",
                    data=f.read(),
                    file_name=f"АПП_Все_{current_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            st.success("Файл успешно создан!")
            
        except Exception as e:
            st.error(f"Ошибка: {e}")
        finally:
            try:
                if os.path.exists(path):
                    os.unlink(path)
            except:
                pass











































































































































































































































































































































































































































































































































































































































