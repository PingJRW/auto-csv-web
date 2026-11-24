import io
import re
from datetime import datetime
import streamlit as st
import pandas as pd
from openpyxl.utils import get_column_letter

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบจัดการไฟล์ CSV อัตโนมัติ", page_icon="📦", layout="centered")
st.title("📦 ระบบจัดการไฟล์ CSV อัตโนมัติ")

# แบ่งหน้าจอเป็น 2 แท็บตามประเภทงาน
tab1, tab2 = st.tabs(["📑 1. จัดการข้อมูลสินค้า (_Processed / _规格 / _证件)", "💰 2. จัดการบิลขนส่งรายวัน (导出财务)"])

# ==========================================
# TAB 1: ระบบจัดการข้อมูลสินค้า
# ==========================================
with tab1:
    st.subheader("อัปโหลดไฟล์สินค้า (สร้าง 3 ไฟล์พร้อมกัน)")
    uploaded_file_1 = st.file_uploader("ลากไฟล์ CSV สินค้ามาวางที่นี่", type=['csv'], key="uploader_tab1")

    if uploaded_file_1 is not None:
        try:
            df = pd.read_csv(
                uploaded_file_1,
                encoding='gb18030',
                dtype={'编码': str, '运单号': str, '收件人电话': str, '收件人身份证号码': str}
            )

            def clean_text(val):
                if pd.isna(val):
                    return ""
                s = str(val)
                if s.endswith('.0'):
                    s = s[:-2]
                s = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', s)
                return s.strip()

            # --- เตรียมข้อมูลสำหรับไฟล์ _证件 ---
            cert_cols = ['用户昵称', '运单号', '申报人名称', '收件人身份证号码']
            existing_cert_cols = [c for c in cert_cols if c in df.columns]
            df_cert = df[existing_cert_cols].copy()

            for col in existing_cert_cols:
                df_cert[col] = df_cert[col].apply(clean_text)

            if '运单号' in df_cert.columns:
                df_cert = df_cert.drop_duplicates(subset=['运单号'], keep='first')

            output3 = io.BytesIO()
            with pd.ExcelWriter(output3, engine='openpyxl') as writer:
                df_cert.to_excel(writer, index=False)
            output3.seek(0)

            # --- ทำความสะอาดคอลัมน์หลัก ---
            cols_to_clean = ['运单号', '编码', '收件人身份证号码', '收件人电话', '申报人名称', '用户昵称']
            for col in cols_to_clean:
                if col in df.columns:
                    df[col] = df[col].apply(clean_text)

            ignore_keywords = ['bento', 'kohkae', 'taokaenoi', 'bigsheet', 'taro', '燕窝', '大哥']

            def process_spec_row(row):
                spec = str(row.get('规格', ''))
                code = str(row.get('编码', '')) if not pd.isna(row.get('编码')) else ''
                prod = str(row.get('品名', '')).lower() if not pd.isna(row.get('品名')) else ''

                if not spec or pd.isna(row.get('规格')) or spec.strip() == '':
                    return pd.Series(["", "", "", False])

                spec = re.sub(r'(?i)f码', '均码', spec)
                spec = spec.lower()

                spec = re.sub(r'[\u4e00-\u9fa5]', lambda m: m.group() if m.group() in '粒片码均' else '', spec)
                if '粒' in spec and '片' in spec:
                    spec = spec.replace('片', '')

                spec = spec.replace(' ', '').replace('/', '')
                spec = re.sub(r'\*1(?!\d)', '', spec)

                if 'g' in spec or 'ml' in spec:
                    spec = spec.replace('粒', '').replace('片', '')
                if 'cm' in spec:
                    spec = spec.replace('片', '')

                if code.startswith('09') and '粒' in spec:
                    spec = spec.replace('粒', 'ml')

                def repl_decimal_g_ml(m):
                    v = float(m.group(1))
                    u = m.group(2)
                    return f"1{u}" if v < 1 else f"{int(v)}{u}"
                spec = re.sub(r'(\d*\.\d+)\s*(g|ml)', repl_decimal_g_ml, spec)

                def repl_decimal_cm(m):
                    v = float(m.group(1))
                    v_mm = v * 10
                    return f"{int(v_mm)}mm" if v_mm.is_integer() else f"{v_mm}mm"
                spec = re.sub(r'(\d*\.\d+)\s*cm', repl_decimal_cm, spec)

                final_spec = spec
                highlight = False

                match = re.search(r'(\d+)\s*(g|ml|粒)\s*\*\s*(\d+)', final_spec)
                if match and (code == '01019900' or code.startswith('09')):
                    base_val = int(match.group(1))
                    unit = match.group(2)
                    multiplier = int(match.group(3))

                    do_multiply = True

                    if base_val > 100:
                        do_multiply = False
                        highlight = True
                    if code in ['09020240', '09010200']:
                        do_multiply = False
                    if '面膜' in prod or '眼膜' in prod:
                        do_multiply = False
                    if code == '01019900':
                        if unit == 'ml':
                            do_multiply = False
                        if any(kw in prod for kw in ignore_keywords):
                            do_multiply = False

                    if do_multiply:
                        new_val = base_val * multiplier
                        final_spec = final_spec[:match.start()] + f"{new_val}{unit}" + final_spec[match.end():]

                base_new_spec = final_spec
                base_new_spec = re.sub(r'\+.*', '', base_new_spec)
                base_new_spec = re.sub(r'\*.*', '', base_new_spec)
                base_new_spec = base_new_spec.strip()

                spec_1 = base_new_spec
                spec_2 = (
                    base_new_spec
                    .replace('均码', 'Free size')
                    .replace('片', ' Pcs')
                    .replace('粒', ' Capsule')
                    .replace('码', ' Size')
                )

                return pd.Series([final_spec, spec_1, spec_2, highlight])

            if '规格' in df.columns:
                spec_results = df.apply(process_spec_row, axis=1)
                df['规格'] = spec_results[0]
                df['_highlight_spec'] = spec_results[3]
            else:
                df['_highlight_spec'] = False

            def get_city(address):
                if pd.isna(address) or str(address).strip() == "":
                    return ""
                addr_str = str(address).replace("中国", "")
                match = re.search(r'(?:省|自治区|^)([^省区]+?)市', addr_str)
                if match:
                    return match.group(1).strip()
                for city in ["北京", "上海", "天津", "重庆"]:
                    if city in addr_str:
                        return city
                return addr_str[:2]

            if '收件人城市' in df.columns and '收件人地址' in df.columns:
                empty_city_mask = df['收件人城市'].isna() | (df['收件人城市'].str.strip() == '')
                if empty_city_mask.any():
                    df.loc[empty_city_mask, '收件人城市'] = df.loc[empty_city_mask, '收件人地址'].apply(get_city)

            if '运单号' in df.columns:
                df = df.sort_values(by='运单号')

            highlight_mask = df['_highlight_spec'].values
            base_name = uploaded_file_1.name.rsplit('.', 1)[0]

            # ไฟล์ที่ 1: _Processed
            df1_save = df.drop(columns=['_highlight_spec'])
            output1 = io.BytesIO()
            with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                styled_df1 = df1_save.style.apply(
                    lambda col: ['background-color: #FFFF00' if m else '' for m in highlight_mask],
                    subset=['规格']
                )
                styled_df1.to_excel(writer, index=False)
            output1.seek(0)

            # ไฟล์ที่ 2: _规格
            df2_save = df.drop(columns=['_highlight_spec'])
            if '规格' in df.columns:
                idx = df2_save.columns.get_loc('规格')
                df2_save.insert(idx, '规格_1', spec_results[1])
                df2_save.insert(idx + 1, '规格_2', spec_results[2])

            output2 = io.BytesIO()
            with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                styled_df_spec = df2_save.style.apply(
                    lambda col: ['background-color: #FFFF00' if m else '' for m in highlight_mask],
                    subset=['规格']
                )
                styled_df_spec.to_excel(writer, index=False)
            output2.seek(0)

            st.success("✅ ประมวลผลสำเร็จ! กดปุ่มด้านล่างเพื่อดาวน์โหลดไฟล์")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button("📥 ดาวน์โหลด _Processed", output1, f"{base_name}_Processed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                st.download_button("📥 ดาวน์โหลด _规格", output2, f"{base_name}_规格.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col3:
                st.download_button("📥 ดาวน์โหลด _证件", output3, f"{base_name}_证件.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")


# ==========================================
# TAB 2: ระบบจัดการบิลขนส่งรายวัน (process_daily_bill.py)
# ==========================================
with tab2:
    st.subheader("อัปโหลดไฟล์บิลรายวัน (สร้างไฟล์ 导出财务YYMMDD.xlsx)")
    uploaded_file_2 = st.file_uploader("ลากไฟล์ CSV บิลรายวันมาวางที่นี่", type=['csv'], key="uploader_tab2")

    VALID_PROVINCES = {
        '北京', '天津', '上海', '重庆',
        '河北省', '山西省', '辽宁省', '吉林省', '黑龙江省',
        '江苏省', '浙江省', '安徽省', '福建省', '江西省', '山东省',
        '河南省', '湖北省', '湖南省', '广东省', '海南省',
        '四川省', '贵州省', '云南省', '陕西省', '甘肃省', '青海省', '台湾省',
        '内蒙古自治区', '广西壮族自治区', '西藏自治区', '宁夏回族自治区', '新疆维吾尔自治区',
        '香港特别行政区', '澳门特别行政区'
    }

    def read_uploaded_csv_auto(file_obj) -> pd.DataFrame:
        raw_bytes = file_obj.getvalue()
        for enc in ['gb18030', 'gbk', 'utf-8-sig', 'utf-8']:
            try:
                df_temp = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, dtype=str)
                df_temp.columns = df_temp.columns.str.strip()
                return df_temp
            except UnicodeDecodeError:
                continue
        raise ValueError("ไม่สามารถอ่านไฟล์ CSV ได้ (รูปแบบ Encoding ไม่รองรับ)")

    def clean_control_chars(series: pd.Series) -> pd.Series:
        return series.fillna('').astype(str).apply(lambda x: re.sub(r'[\x00-\x1f\x7f]', '', x))

    def pinyin_sort_key(series: pd.Series) -> pd.Series:
        return series.fillna('').astype(str).apply(
            lambda x: x.lower().encode('gb18030', errors='replace')
        )

    def calc_display_width(text: str) -> int:
        return sum(2 if ord(ch) > 127 else 1 for ch in str(text))

    if uploaded_file_2 is not None:
        try:
            df_bill = read_uploaded_csv_auto(uploaded_file_2)

            target_cols = ['用户昵称', '运单号', '收件人名称', '省', '毛重', '长度', '宽度', '高度', '税金']
            missing_cols = [c for c in target_cols if c not in df_bill.columns]
            if missing_cols:
                st.error(f"❌ ไม่พบคอลัมน์ที่จำเป็นในไฟล์: {', '.join(missing_cols)}")
                st.stop()

            df_bill = df_bill[target_cols].copy()

            # Clean อักขระควบคุมทุกคอลัมน์
            for col in target_cols:
                df_bill[col] = clean_control_chars(df_bill[col]).str.strip()

            # แปลงตัวเลข
            int_cols = ['毛重', '长度', '宽度', '高度']
            for col in int_cols:
                df_bill[col] = pd.to_numeric(df_bill[col], errors='coerce').fillna(0).astype(int)

            df_bill['税金'] = pd.to_numeric(df_bill['税金'], errors='coerce').fillna(0.0)

            # ลบแถวที่ 运单号 ซ้ำกัน
            df_bill = df_bill.drop_duplicates(subset=['运单号'], keep='first').copy()

            # ตรวจคอลัมน์ 省
            df_bill['省'] = (
                df_bill['省']
                .str.replace(' ', '', regex=False)
                .str.replace('\u3000', '', regex=False)
                .str.replace('市', '', regex=False)
            )

            invalid_rows = df_bill[~df_bill['省'].isin(VALID_PROVINCES)]
            warnings = [
                f"- 运单号: `{row['运单号']}` | ค่าในคอลัมน์ 省: `'{row['省']}'`"
                for _, row in invalid_rows.iterrows()
            ]

            # หารหน่วยตามเงื่อนไข
            df_bill['毛重'] = df_bill['毛重'] / 1000.0
            df_bill['长度'] = df_bill['长度'] / 10.0
            df_bill['宽度'] = df_bill['宽度'] / 10.0
            df_bill['高度'] = df_bill['高度'] / 10.0

            # เรียงลำดับข้อมูลตามพินอิน
            df_bill = df_bill.sort_values(
                by='用户昵称',
                ascending=True,
                kind='mergesort',
                key=pinyin_sort_key
            ).reset_index(drop=True)

            # สร้างไฟล์ Excel ใน Memory (BytesIO)
            yymmdd = datetime.now().strftime('%y%m%d')
            bill_filename = f"导出财务{yymmdd}.xlsx"

            output_bill = io.BytesIO()
            with pd.ExcelWriter(output_bill, engine='openpyxl') as writer:
                df_bill.to_excel(writer, index=False, sheet_name='Sheet1')
                ws = writer.sheets['Sheet1']

                # บังคับคอลัมน์ 运单号 เป็น Text (@)
                for row_idx in range(2, len(df_bill) + 2):
                    cell = ws.cell(row=row_idx, column=2)
                    cell.number_format = '@'

                # ปรับความกว้างคอลัมน์อัตโนมัติ
                for col_idx, col_name in enumerate(df_bill.columns, start=1):
                    col_letter = get_column_letter(col_idx)
                    max_data_width = df_bill[col_name].map(calc_display_width).max() if not df_bill.empty else 0
                    max_len = max(max_data_width, calc_display_width(col_name))
                    ws.column_dimensions[col_letter].width = max_len + 4

            output_bill.seek(0)

            # แสดงผลลัพธ์และแจ้งเตือนบนหน้าเว็บ
            st.success(f"✅ จัดระเบียบข้อมูลสำเร็จ! จำนวนรายการสุทธิ: **{len(df_bill)}** แถว")

            if warnings:
                st.warning(f"⚠️ พบชื่อมณฑล (省) ที่ไม่ตรงกับรายชื่อมาตรฐาน **{len(warnings)}** รายการ:")
                with st.expander("คลิกเพื่อดูรายการที่ต้องตรวจสอบ", expanded=True):
                    st.markdown("\n".join(warnings))
            else:
                st.info("✅ ตรวจสอบชื่อมณฑล (省) ถูกต้องครบทุกรายการ")

            st.download_button(
                label=f"📥 ดาวน์โหลดไฟล์ {bill_filename}",
                data=output_bill,
                file_name=bill_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลบิล: {e}")
