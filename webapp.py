import streamlit as st
import pandas as pd
import re
import io

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบจัดการไฟล์ CSV", page_icon="📦")
st.title("📦 ระบบจัดการไฟล์ CSV อัตโนมัติ")
st.write("อัปโหลดไฟล์ต้นฉบับ (GB18030) ระบบจะจัดการข้อมูลและสร้างไฟล์ Excel ให้ทันที")

# สร้างกล่องอัปโหลดไฟล์
uploaded_file = st.file_uploader("ลากไฟล์ CSV มาวางที่นี่ หรือกดปุ่ม Browse files", type=['csv'])

if uploaded_file is not None:
    try:
        # อ่านไฟล์ที่ผู้ใช้อัปโหลด
        df = pd.read_csv(uploaded_file, encoding='gb18030', dtype={'编码': str, '运单号': str, '收件人电话': str, '收件人身份证号码': str})
        
        # --- เริ่มกระบวนการทำความสะอาดข้อมูล ---
        def clean_text(val):
            if pd.isna(val): return ""
            s = str(val)
            if s.endswith('.0'): s = s[:-2]
            s = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', s)
            return s.strip()

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
            spec = spec.replace('/', '')
            
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
                if code in ['09020240', '09010200']: do_multiply = False
                if '面膜' in prod or '眼膜' in prod: do_multiply = False
                if code == '01019900':
                    if unit == 'ml': do_multiply = False
                    if any(kw in prod for kw in ignore_keywords): do_multiply = False
                    
                if do_multiply:
                    new_val = base_val * multiplier
                    final_spec = final_spec[:match.start()] + f"{new_val}{unit}" + final_spec[match.end():]

            base_new_spec = final_spec 
            base_new_spec = re.sub(r'\+.*', '', base_new_spec)
            base_new_spec = re.sub(r'\*.*', '', base_new_spec)
            base_new_spec = base_new_spec.strip()
            
            spec_1 = base_new_spec
            spec_2 = base_new_spec.replace('片', ' Pcs').replace('粒', ' Capsule').replace('码', ' Size')
            
            return pd.Series([final_spec, spec_1, spec_2, highlight])

        if '规格' in df.columns:
            spec_results = df.apply(process_spec_row, axis=1)
            df['规格'] = spec_results[0]
            df['_highlight_spec'] = spec_results[3]
        else:
            df['_highlight_spec'] = False

        def get_city(address):
            if pd.isna(address) or str(address).strip() == "": return ""
            addr_str = str(address).replace("中国", "")
            match = re.search(r'(?:省|自治区|^)([^省区]+?)市', addr_str)
            if match: return match.group(1).strip()
            for city in ["北京", "上海", "天津", "重庆"]:
                if city in addr_str: return city
            return addr_str[:2]

        empty_city_mask = df['收件人城市'].isna() | (df['收件人城市'].str.strip() == '')
        if empty_city_mask.any():
            df.loc[empty_city_mask, '收件人城市'] = df.loc[empty_city_mask, '收件人地址'].apply(get_city)

        if '运单号' in df.columns:
            df = df.sort_values(by='运单号')

        # --- เตรียมไฟล์ให้พร้อมดาวน์โหลด ---
        highlight_mask = df['_highlight_spec'].values
        base_name = uploaded_file.name.rsplit('.', 1)[0]
        
        # ไฟล์ที่ 1: _Processed
        df1_save = df.drop(columns=['_highlight_spec'])
        output1 = io.BytesIO()
        with pd.ExcelWriter(output1, engine='openpyxl') as writer:
            styled_df1 = df1_save.style.apply(lambda col: ['background-color: #FFFF00' if m else '' for m in highlight_mask], subset=['规格'])
            styled_df1.to_excel(writer, index=False)
        output1.seek(0)
        
        # ไฟล์ที่ 2: _规格
        df2_save = df.drop(columns=['_highlight_spec'])
        if '规格' in df.columns:
            idx = df2_save.columns.get_loc('规格')
            df2_save.insert(idx, '规格_1', spec_results[1])
            df2_save.insert(idx+1, '规格_2', spec_results[2])
        
        output2 = io.BytesIO()
        with pd.ExcelWriter(output2, engine='openpyxl') as writer:
            styled_df_spec = df2_save.style.apply(lambda col: ['background-color: #FFFF00' if m else '' for m in highlight_mask], subset=['规格'])
            styled_df_spec.to_excel(writer, index=False)
        output2.seek(0)
        
        # แสดงผลบนหน้าเว็บ
        st.success("✅ ประมวลผลสำเร็จ! กดปุ่มด้านล่างเพื่อดาวน์โหลดไฟล์")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(label="📥 ดาวน์โหลด _Processed", data=output1, file_name=f"{base_name}_Processed.xlsx", mime="application/vnd.ms-excel")
        with col2:
            st.download_button(label="📥 ดาวน์โหลด _规格", data=output2, file_name=f"{base_name}_规格.xlsx", mime="application/vnd.ms-excel")
            
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")