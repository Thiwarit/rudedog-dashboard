import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Dashboard สต็อก Rudedog", layout="wide")
st.title("📦 Dashboard สต็อก – Rudedog (V3 Fixed - No Plotly)")

uploaded_file = st.file_uploader("📁 อัปโหลดไฟล์ Excel (6 ชีต)", type=["xlsx"])

def normalize_data(df):
    """Normalize data without changing column names"""
    df_copy = df.copy()
    # Only normalize the รหัสรูปแบบสินค้า column for matching
    if "รหัสรูปแบบ" in df_copy.columns:
        df_copy["รหัสรูปแบบ_normalized"] = df_copy["รหัสรูปแบบ"].astype(str).str.strip().str.lower()
    elif "รหัสรูปแบบสินค้า" in df_copy.columns:
        df_copy["รหัสรูปแบบสินค้า_normalized"] = df_copy["รหัสรูปแบบสินค้า"].astype(str).str.strip().str.lower()
    return df_copy

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.info(f"ชีตที่พบในไฟล์: {sheet_names}")
        
        # Read main data - try different possible sheet names
        main_sheet_name = None
        for name in ["Sheet1", "Stock", "stock", "Item 21568", "Main", "Data"]:
            if name in sheet_names:
                main_sheet_name = name
                break
        
        if not main_sheet_name:
            st.error("ไม่พบ Sheet หลัก กรุณาตรวจสอบชื่อ Sheet")
            st.stop()
        
        # Read all sheets
        df_main = pd.read_excel(uploaded_file, sheet_name=main_sheet_name)
        df_price = pd.read_excel(uploaded_file, sheet_name="ราคา") if "ราคา" in sheet_names else pd.DataFrame()
        df_sm = pd.read_excel(uploaded_file, sheet_name="SM") if "SM" in sheet_names else pd.DataFrame()
        df_exclude = pd.read_excel(uploaded_file, sheet_name="ตัดออก") if "ตัดออก" in sheet_names else pd.DataFrame()
        df_fg_market = pd.read_excel(uploaded_file, sheet_name="FG รุ่นทำตลาด") if "FG รุ่นทำตลาด" in sheet_names else pd.DataFrame()
        df_sm_active = pd.read_excel(uploaded_file, sheet_name="SM ใช้งาน") if "SM ใช้งาน" in sheet_names else pd.DataFrame()
        
        # Debug: Show column names
        with st.expander("🔍 ตรวจสอบ Column Names"):
            st.write("Main sheet columns:", df_main.columns.tolist())
            if not df_price.empty:
                st.write("Price sheet columns:", df_price.columns.tolist())
        
        # Clean main data
        df_all = df_main.copy()
        
        # Handle different possible column names
        stock_col = None
        for col in ["stock", "Stock", "จํานวนที่ใช้ได้", "จำนวนที่ใช้ได้"]:
            if col in df_all.columns:
                stock_col = col
                break
        
        wip_col = None
        for col in ["อยู่ระหว่างการจัดซื้อ", "WIP", "จัดซื้อ"]:
            if col in df_all.columns:
                wip_col = col
                break
        
        code_col = None
        for col in ["รหัสรูปแบบ", "รหัสรูปแบบสินค้า", "รหัส"]:
            if col in df_all.columns:
                code_col = col
                break
        
        if not stock_col or not code_col:
            st.error("ไม่พบ Column ที่จำเป็น กรุณาตรวจสอบ Column Names")
            st.stop()
        
        # Convert to numeric
        df_all[stock_col] = pd.to_numeric(df_all[stock_col], errors="coerce").fillna(0)
        if wip_col:
            df_all[wip_col] = pd.to_numeric(df_all[wip_col], errors="coerce").fillna(0)
        
        # Create price mapping
        price_map = {}
        if not df_price.empty:
            price_code_col = None
            for col in ["รหัสรูปแบบ", "รหัสรูปแบบสินค้า", "รหัส"]:
                if col in df_price.columns:
                    price_code_col = col
                    break
            
            price_value_col = None
            for col in ["ราคาต่อชิ้น", "ราคา", "price"]:
                if col in df_price.columns:
                    price_value_col = col
                    break
            
            if price_code_col and price_value_col:
                df_price[price_code_col] = df_price[price_code_col].astype(str).str.strip().str.lower()
                price_map = df_price.set_index(price_code_col)[price_value_col].to_dict()
        
        # Create prefix price mapping for FG
        prefix_price_map = {}
        for k, v in price_map.items():
            if "-" in k:
                prefix = k.split("-")[0]
                if prefix not in prefix_price_map:
                    prefix_price_map[prefix] = v
        
        def assign_price(sku):
            if pd.isna(sku):
                return 0
            sku_lower = str(sku).strip().lower()
            
            # 1. หาราคาแบบตรงตัว
            if sku_lower in price_map:
                return price_map[sku_lower]
            
            # 2. หาราคาจาก prefix (TS-Dogblur -> TS)
            if "-" in sku_lower:
                prefix = sku_lower.split("-")[0]
                if prefix in price_map:
                    return price_map[prefix]
            
            # 3. ถ้าหาไม่เจอเลย ให้ราคา default
            return 100
        
        def get_price_source(sku):
            """ตรวจสอบว่าราคามาจากไหน สำหรับ debug"""
            if pd.isna(sku):
                return "N/A"
            sku_lower = str(sku).strip().lower()
            
            if sku_lower in price_map:
                return "ตรงตัว"
            elif "-" in sku_lower and sku_lower.split("-")[0] in price_map:
                return f"Prefix ({sku_lower.split('-')[0]})"
            else:
                return "Default (100)"
        
        df_all["ราคาต่อชิ้น"] = df_all[code_col].apply(assign_price)
        df_all["แหล่งราคา"] = df_all[code_col].apply(get_price_source)
        df_all["มูลค่า"] = df_all[stock_col] * df_all["ราคาต่อชิ้น"]
        if wip_col:
            df_all["มูลค่า_WIP"] = df_all[wip_col] * df_all["ราคาต่อชิ้น"]
        
        # Find products without proper pricing (using default price) that have stock or sales
        sales_columns = ['ยอดขาย3วัน', 'ยอดขาย 7', 'ยอดขาย 15', 'ยอดขาย 30', 'ยอดขาย 60 วัน', 'ยอดขาย 90 วัน']
        existing_sales_cols = [col for col in sales_columns if col in df_all.columns]
        
        # Calculate total sales for each product
        if existing_sales_cols:
            df_all['ยอดขายรวม'] = df_all[existing_sales_cols].sum(axis=1)
        else:
            df_all['ยอดขายรวม'] = 0
        
        # Products with stock or sales but no proper price
        no_price_products = df_all[
            (df_all["แหล่งราคา"] == "Default (100)") & 
            ((df_all[stock_col] > 0) | (df_all['ยอดขายรวม'] > 0))
        ].copy()
        
        # Sort by highest stock + sales first
        if len(no_price_products) > 0:
            no_price_products['คะแนนความสำคัญ'] = no_price_products[stock_col] + no_price_products['ยอดขายรวม']
            no_price_products = no_price_products.sort_values('คะแนนความสำคัญ', ascending=False)
        
        # Create sets for filtering
        sm_set = set()
        if not df_sm.empty:
            sm_col = None
            for col in ["รหัสรูปแบบ", "รหัสรูปแบบสินค้า", "รหัส"]:
                if col in df_sm.columns:
                    sm_col = col
                    break
            if sm_col:
                sm_set = set(df_sm[sm_col].astype(str).str.strip().str.lower())
        
        exclude_set = set()
        if not df_exclude.empty:
            exclude_col = None
            for col in ["รหัสรูปแบบ", "รหัสรูปแบบสินค้า", "รหัส"]:
                if col in df_exclude.columns:
                    exclude_col = col
                    break
            if exclude_col:
                exclude_set = set(df_exclude[exclude_col].astype(str).str.strip().str.lower())
        
        fg_market_set = set()
        if not df_fg_market.empty:
            fg_market_col = None
            for col in ["รหัสรูปแบบ", "รหัสรูปแบบสินค้า", "รหัส"]:
                if col in df_fg_market.columns:
                    fg_market_col = col
                    break
            if fg_market_col:
                fg_market_set = set(df_fg_market[fg_market_col].astype(str).str.strip().str.lower())
        
        sm_active_set = set()
        if not df_sm_active.empty:
            sm_active_col = None
            for col in ["รหัสรูปแบบ", "รหัสรูปแบบสินค้า", "รหัส"]:
                if col in df_sm_active.columns:
                    sm_active_col = col
                    break
            if sm_active_col:
                sm_active_set = set(df_sm_active[sm_active_col].astype(str).str.strip().str.lower())
        
        # Create normalized column for filtering
        df_all["รหัสรูปแบบ_normalized"] = df_all[code_col].astype(str).str.strip().str.lower()
        
        # Filter conditions
        zip_cut = df_all["รหัสรูปแบบ_normalized"].str.startswith("zip")
        is_number_only = df_all["รหัสรูปแบบ_normalized"].str.fullmatch(r"\d+")
        
        # Create filtered dataframes with corrected logic
        # 1. SM ทั้งหมด - รหัสที่อยู่ในชีต SM
        df_sm_all = df_all[df_all["รหัสรูปแบบ_normalized"].isin(sm_set)]
        
        # 2. FG ทั้งหมด - ที่ไม่อยู่ใน SM, ไม่อยู่ใน "ตัดออก", ไม่ขึ้นต้น zip, และไม่ใช่เลขล้วน
        df_fg_all = df_all[
            ~df_all["รหัสรูปแบบ_normalized"].isin(sm_set) & 
            ~df_all["รหัสรูปแบบ_normalized"].isin(exclude_set) & 
            ~zip_cut & 
            ~is_number_only
        ]
        
        # 3. FG รุ่นทำตลาด - FG ที่อยู่ในชีต "FG รุ่นทำตลาด" (จากข้อมูลทั้งหมด ไม่จำกัดแค่ FG)
        df_fg_market_filtered = df_all[df_all["รหัสรูปแบบ_normalized"].isin(fg_market_set)]
        
        # 4. SM ใช้งาน - รหัสในชีต "SM ใช้งาน" (จากข้อมูลทั้งหมด ไม่จำกัดแค่ SM)
        df_sm_use = df_all[df_all["รหัสรูปแบบ_normalized"].isin(sm_active_set)]
        
        # Calculate totals for dashboard
        # SM stock + SM WIP + FG stock + FG WIP
        sm_stock_total = df_sm_all[stock_col].sum()
        sm_wip_total = df_sm_all[wip_col].sum() if wip_col else 0
        fg_stock_total = df_fg_all[stock_col].sum()
        fg_wip_total = df_fg_all[wip_col].sum() if wip_col else 0
        
        # Total items calculation
        total_items_count = sm_stock_total + sm_wip_total + fg_stock_total + fg_wip_total
        
        # Value calculations
        sm_stock_value = (df_sm_all[stock_col] * df_sm_all["ราคาต่อชิ้น"]).sum()
        sm_wip_value = (df_sm_all[wip_col] * df_sm_all["ราคาต่อชิ้น"]).sum() if wip_col else 0
        fg_stock_value = (df_fg_all[stock_col] * df_fg_all["ราคาต่อชิ้น"]).sum()
        fg_wip_value = (df_fg_all[wip_col] * df_fg_all["ราคาต่อชิ้น"]).sum() if wip_col else 0
        
        total_value = sm_stock_value + sm_wip_value + fg_stock_value + fg_wip_value
        
        # Debug information
        with st.expander("🔍 ตรวจสอบการจัดกลุ่มและการคำนวณ"):
            st.write(f"**ชีตหลักที่ใช้:** {main_sheet_name}")
            st.write(f"**Column สต็อก:** {stock_col}")
            st.write(f"**Column WIP:** {wip_col}")
            st.write("")
            
            # Price mapping summary
            st.write("**การหาราคา:**")
            st.write(f"- จำนวนราคาในชีต 'ราคา': {len(price_map)} รายการ")
            if len(price_map) > 0:
                st.write(f"- ตัวอย่างราคา: {dict(list(price_map.items())[:5])}")
            
            # Price source breakdown
            price_source_count = df_all['แหล่งราคา'].value_counts()
            st.write("**แหล่งที่มาของราคา:**")
            for source, count in price_source_count.items():
                st.write(f"- {source}: {count:,} รายการ")
            
            st.write("")
            st.write(f"**SM ทั้งหมด:** {len(df_sm_all)} รายการ (รหัสที่อยู่ในชีต SM)")
            st.write(f"  - SM Stock: {sm_stock_total:,.0f} ชิ้น = {sm_stock_value:,.0f} บาท")
            if wip_col:
                st.write(f"  - SM WIP: {sm_wip_total:,.0f} ชิ้น = {sm_wip_value:,.0f} บาท")
            st.write("")
            st.write(f"**FG ทั้งหมด:** {len(df_fg_all)} รายการ (ไม่อยู่ใน SM, ไม่อยู่ใน ตัดออก, ไม่ขึ้นต้น zip, ไม่ใช่เลขล้วน)")
            st.write(f"  - FG Stock: {fg_stock_total:,.0f} ชิ้น = {fg_stock_value:,.0f} บาท")
            if wip_col:
                st.write(f"  - FG WIP: {fg_wip_total:,.0f} ชิ้น = {fg_wip_value:,.0f} บาท")
            st.write("")
            st.write(f"**FG รุ่นทำตลาด:** {len(df_fg_market_filtered)} รายการ (รหัสที่อยู่ในชีต FG รุ่นทำตลาด)")
            st.write(f"**SM ใช้งาน:** {len(df_sm_use)} รายการ (รหัสที่อยู่ในชีต SM ใช้งาน)")
            st.write("")
            st.write(f"**รวมข้อมูลทั้งหมดในไฟล์:** {len(df_all)} แถว")
            st.write(f"**รายการทั้งหมดที่นับ:** {total_items_count:,.0f} ชิ้น (SM Stock + SM WIP + FG Stock + FG WIP)")
            st.write(f"**มูลค่ารวม:** {total_value:,.0f} บาท")
            
            # Show sample data from each set
            if len(sm_set) > 0:
                st.write(f"ตัวอย่าง SM set: {list(sm_set)[:5]}")
            if len(fg_market_set) > 0:
                st.write(f"ตัวอย่าง FG รุ่นทำตลาด set: {list(fg_market_set)[:5]}")
            if len(sm_active_set) > 0:
                st.write(f"ตัวอย่าง SM ใช้งาน set: {list(sm_active_set)[:5]}")
            if len(exclude_set) > 0:
                st.write(f"ตัวอย่าง ตัดออก set: {list(exclude_set)[:5]}")
        
        def sales_sum(df, col):
            return df[col].sum() if col in df.columns else 0
        
        def calculate_period_sales(df, col1, col2):
            """Calculate sales for a specific period (col1 - col2)"""
            if col1 in df.columns and col2 in df.columns:
                return (df[col1] - df[col2]).sum()
            elif col1 in df.columns:
                return df[col1].sum()
            return 0
        
        # Dashboard Display
        st.header("📊 สรุปข้อมูลหลัก")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("รายการทั้งหมด", f"{total_items_count:,.0f}")
            st.caption("SM + SM WIP + FG + FG WIP")
            st.metric("มูลค่ารวม", f"{total_value:,.0f} บาท")
        
        with col2:
            st.metric("SM จำนวน", f"{sm_stock_total:,.0f}")
            st.metric("SM มูลค่า", f"{sm_stock_value:,.0f} บาท")
            if wip_col:
                st.metric("SM WIP", f"{sm_wip_total:,.0f}", f"{sm_wip_value:,.0f} บาท")
        
        with col3:
            st.metric("FG จำนวน", f"{fg_stock_total:,.0f}")
            st.metric("FG มูลค่า", f"{fg_stock_value:,.0f} บาท")
            if wip_col:
                st.metric("FG WIP", f"{fg_wip_total:,.0f}", f"{fg_wip_value:,.0f} บาท")
        
        with col4:
            # Summary breakdown
            st.metric("รวม SM ทั้งหมด", f"{sm_stock_total + sm_wip_total:,.0f}")
            st.caption("Stock + WIP")
            st.metric("รวม FG ทั้งหมด", f"{fg_stock_total + fg_wip_total:,.0f}")
            st.caption("Stock + WIP")
        
        st.divider()
        
        # Warning for products without proper pricing
        if len(no_price_products) > 0:
            st.warning(f"⚠️ **พบสินค้าที่ไม่มีราคาในชีต 'ราคา' จำนวน {len(no_price_products)} รายการ**")
            st.write("สินค้าเหล่านี้ใช้ราคา Default = 100 บาท ซึ่งอาจทำให้การคำนวณมูลค่าไม่ถูกต้อง")
            
            with st.expander("📋 ดูรายการสินค้าที่ไม่มีราคา (เรียงตามความสำคัญ)"):
                display_cols = [code_col, 'ชื่อสินค้า', stock_col, 'ยอดขายรวม', 'ราคาต่อชิ้น', 'แหล่งราคา']
                if wip_col:
                    display_cols.insert(-2, wip_col)
                
                available_cols = [col for col in display_cols if col in no_price_products.columns]
                st.dataframe(no_price_products[available_cols].head(20), use_container_width=True)
                
                if len(no_price_products) > 20:
                    st.info(f"แสดง 20 รายการแรก จากทั้งหมด {len(no_price_products)} รายการ")
                
                # Summary of missing prices
                st.subheader("📊 สรุปสินค้าที่ไม่มีราคา")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("จำนวนรายการ", f"{len(no_price_products):,}")
                
                with col2:
                    missing_stock = no_price_products[stock_col].sum()
                    st.metric("สต็อกรวม", f"{missing_stock:,.0f}")
                
                with col3:
                    missing_sales = no_price_products['ยอดขายรวม'].sum()
                    st.metric("ยอดขายรวม", f"{missing_sales:,.0f}")
        
        # Add bar charts
        st.divider()
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 กราฟจำนวน")
            # Create data for quantity chart
            quantity_data = pd.DataFrame({
                'หมวดหมู่': ['SM', 'SM WIP', 'FG', 'FG WIP'],
                'จำนวน': [sm_stock_total, sm_wip_total, fg_stock_total, fg_wip_total]
            }).set_index('หมวดหมู่')
            
            st.bar_chart(quantity_data)
            
            # Show values as table
            st.write("**ข้อมูลจำนวน:**")
            quantity_display = pd.DataFrame({
                'หมวดหมู่': ['SM', 'SM WIP', 'FG', 'FG WIP'],
                'จำนวน': [f"{sm_stock_total:,.0f}", f"{sm_wip_total:,.0f}", 
                          f"{fg_stock_total:,.0f}", f"{fg_wip_total:,.0f}"]
            })
            st.dataframe(quantity_display, use_container_width=True, hide_index=True)
        
        with col2:
            st.subheader("💰 กราฟมูลค่า")
            # Create data for value chart
            value_data = pd.DataFrame({
                'หมวดหมู่': ['SM', 'SM WIP', 'FG', 'FG WIP'],
                'มูลค่า': [sm_stock_value, sm_wip_value, fg_stock_value, fg_wip_value]
            }).set_index('หมวดหมู่')
            
            st.bar_chart(value_data)
            
            # Show values as table
            st.write("**ข้อมูลมูลค่า:**")
            value_display = pd.DataFrame({
                'หมวดหมู่': ['SM', 'SM WIP', 'FG', 'FG WIP'],
                'มูลค่า (บาท)': [f"{sm_stock_value:,.0f}", f"{sm_wip_value:,.0f}", 
                               f"{fg_stock_value:,.0f}", f"{fg_wip_value:,.0f}"]
            })
            st.dataframe(value_display, use_container_width=True, hide_index=True)
        
        # Tabs for detailed analysis
        tab1, tab2, tab3, tab4 = st.tabs(["📊 ภาพรวม", "🔧 SM", "🧥 FG", "📈 ยอดขาย"])
        
        with tab1:
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🔧 SM ที่ใช้งาน")
                if len(df_sm_use) > 0:
                    st.metric("จำนวน", f"{df_sm_use[stock_col].sum():,.0f}")
                    st.metric("มูลค่า", f"{df_sm_use['มูลค่า'].sum():,.0f} บาท")
                    
                    # Show which items are included
                    st.write(f"รวม {len(df_sm_use)} รายการจากชีต 'SM ใช้งาน'")
                    
                    # Simple bar chart using streamlit
                    sm_summary = df_sm_use.groupby(code_col)['มูลค่า'].sum().sort_values(ascending=False)
                    if len(sm_summary) > 0:
                        st.subheader("มูลค่า SM ที่ใช้งาน")
                        st.bar_chart(sm_summary)
                        
                        # Show top 5 as table
                        st.subheader("Top 5 SM ที่ใช้งาน")
                        top_5_sm = sm_summary.head(5).reset_index()
                        top_5_sm.columns = ['รหัสรูปแบบ', 'มูลค่า']
                        st.dataframe(top_5_sm, use_container_width=True)
                else:
                    st.info("ไม่พบข้อมูล SM ที่ใช้งาน หรือชีต 'SM ใช้งาน' ว่างเปล่า")
            
            with col2:
                st.subheader("🧥 FG รุ่นทำตลาด")
                if len(df_fg_market_filtered) > 0:
                    st.metric("จำนวน", f"{df_fg_market_filtered[stock_col].sum():,.0f}")
                    st.metric("มูลค่า", f"{df_fg_market_filtered['มูลค่า'].sum():,.0f} บาท")
                    
                    # Show which items are included
                    st.write(f"รวม {len(df_fg_market_filtered)} รายการจากชีต 'FG รุ่นทำตลาด'")
                    
                    # Simple bar chart using streamlit
                    fg_summary = df_fg_market_filtered.groupby(code_col)['มูลค่า'].sum().sort_values(ascending=False)
                    if len(fg_summary) > 0:
                        st.subheader("มูลค่า FG ทำตลาด (Top 10)")
                        st.bar_chart(fg_summary.head(10))
                        
                        # Show top 5 as table
                        st.subheader("Top 5 FG ทำตลาด")
                        top_5_fg = fg_summary.head(5).reset_index()
                        top_5_fg.columns = ['รหัสรูปแบบ', 'มูลค่า']
                        st.dataframe(top_5_fg, use_container_width=True)
                else:
                    st.info("ไม่พบข้อมูล FG รุ่นทำตลาด หรือชีต 'FG รุ่นทำตลาด' ว่างเปล่า")
        
        with tab2:
            st.subheader("🔧 ข้อมูล SM")
            if len(df_sm_all) > 0:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("จำนวนรายการ", f"{len(df_sm_all):,}")
                with col2:
                    st.metric("สต็อกรวม", f"{df_sm_all[stock_col].sum():,.0f}")
                with col3:
                    st.metric("มูลค่ารวม", f"{df_sm_all['มูลค่า'].sum():,.0f} บาท")
                
                # SM Data Table
                display_cols = [code_col, "ชื่อสินค้า", stock_col, "ราคาต่อชิ้น", "มูลค่า"]
                if wip_col:
                    display_cols.append(wip_col)
                
                available_cols = [col for col in display_cols if col in df_sm_all.columns]
                st.dataframe(df_sm_all[available_cols], use_container_width=True)
                
                # SM Distribution Chart
                if len(df_sm_all) > 0:
                    sm_dist = df_sm_all.groupby(code_col)[stock_col].sum().sort_values(ascending=False)
                    st.subheader("การกระจายสต็อก SM")
                    st.bar_chart(sm_dist)
            else:
                st.info("ไม่พบข้อมูล SM")
        
        with tab3:
            st.subheader("🧥 ข้อมูล FG")
            if len(df_fg_all) > 0:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("จำนวนรายการ", f"{len(df_fg_all):,}")
                with col2:
                    st.metric("สต็อกรวม", f"{df_fg_all[stock_col].sum():,.0f}")
                with col3:
                    st.metric("มูลค่ารวม", f"{df_fg_all['มูลค่า'].sum():,.0f} บาท")
                
                # FG Category Filter
                st.subheader("🔍 กรองตามหมวดหมู่")
                categories = ["ทั้งหมด", "TS", "PL", "LC", "ACS", "อื่นๆ"]
                selected_cat = st.selectbox("เลือกหมวดหมู่", categories)
                
                if selected_cat == "ทั้งหมด":
                    filtered_fg = df_fg_all
                elif selected_cat == "อื่นๆ":
                    filtered_fg = df_fg_all[~df_fg_all["รหัสรูปแบบ_normalized"].str.startswith(("ts-", "pl-", "lc-", "acs-"))]
                else:
                    filtered_fg = df_fg_all[df_fg_all["รหัสรูปแบบ_normalized"].str.startswith(selected_cat.lower() + "-")]
                
                # Display filtered data
                if len(filtered_fg) > 0:
                    st.metric(f"จำนวน {selected_cat}", f"{filtered_fg[stock_col].sum():,.0f}")
                    st.metric(f"มูลค่า {selected_cat}", f"{filtered_fg['มูลค่า'].sum():,.0f} บาท")
                    
                    # Chart for selected category
                    if len(filtered_fg) > 0:
                        fg_cat_dist = filtered_fg.groupby(code_col)['มูลค่า'].sum().sort_values(ascending=False).head(10)
                        st.subheader(f"มูลค่า {selected_cat} (Top 10)")
                        st.bar_chart(fg_cat_dist)
                    
                    display_cols = [code_col, "ชื่อสินค้า", stock_col, "ราคาต่อชิ้น", "มูลค่า"]
                    available_cols = [col for col in display_cols if col in filtered_fg.columns]
                    st.dataframe(filtered_fg[available_cols], use_container_width=True)
                else:
                    st.info(f"ไม่พบข้อมูล FG หมวด {selected_cat}")
            else:
                st.info("ไม่พบข้อมูล FG")
        
        with tab4:
            st.subheader("📈 ยอดขาย FG")
            
            # Sales analysis
            sales_periods = [
                ("ยอดขาย3วัน", "3 วัน"),
                ("ยอดขาย 7", "7 วัน"), 
                ("ยอดขาย 15", "15 วัน"),
                ("ยอดขาย 30", "1 เดือน")
            ]
            
            # Calculate derived periods
            period_1_month_ago = calculate_period_sales(df_fg_all, "ยอดขาย 60 วัน", "ยอดขาย 30")
            period_2_month_ago = calculate_period_sales(df_fg_all, "ยอดขาย 90 วัน", "ยอดขาย 60 วัน")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 ยอดขาย FG ทั้งหมด")
                for col_name, period_name in sales_periods:
                    value = sales_sum(df_fg_all, col_name)
                    st.metric(f"ยอดขาย {period_name}", f"{value:,.0f}")
                
                st.metric("ยอดขาย 1 เดือนก่อน", f"{period_1_month_ago:,.0f}")
                st.metric("ยอดขาย 2 เดือนก่อน", f"{period_2_month_ago:,.0f}")
                
                # Sales trend chart as bar chart with values
                if any(col_name in df_fg_all.columns for col_name, _ in sales_periods):
                    # Create ordered data for proper sequence
                    ordered_periods = [
                        ("ยอดขาย3วัน", "3 วัน"),
                        ("ยอดขาย 7", "7 วัน"), 
                        ("ยอดขาย 15", "15 วัน"),
                        ("ยอดขาย 30", "1 เดือน")
                    ]
                    
                    sales_data = []
                    labels = []
                    
                    # Add main periods in order
                    for col_name, period_name in ordered_periods:
                        if col_name in df_fg_all.columns:
                            sales_data.append(sales_sum(df_fg_all, col_name))
                            labels.append(period_name)
                    
                    # Add derived periods at the end
                    if period_1_month_ago > 0:
                        sales_data.append(period_1_month_ago)
                        labels.append("1 เดือนก่อน")
                    if period_2_month_ago > 0:
                        sales_data.append(period_2_month_ago)
                        labels.append("2 เดือนก่อน")
                    
                    if sales_data:
                        st.subheader("📈 แนวโน้มยอดขาย FG")
                        
                        # Create DataFrame with ordered index using numbers for proper sorting
                        sales_df = pd.DataFrame({
                            'ช่วงเวลา': [f"{i+1:02d}. {label}" for i, label in enumerate(labels)],
                            'ยอดขาย': sales_data
                        })
                        
                        # Display as bar chart with ordered index
                        st.bar_chart(sales_df.set_index('ช่วงเวลา'))
                        
                        # Display values table for reference (without numbers)
                        st.write("**ค่าตัวเลขยอดขาย:**")
                        for i, (label, value) in enumerate(zip(labels, sales_data)):
                            st.write(f"**{label}:** {value:,.0f}")
            
            with col2:
                st.subheader("🎯 ยอดขาย FG รุ่นทำตลาด")
                if len(df_fg_market_filtered) > 0:
                    for col_name, period_name in sales_periods:
                        value = sales_sum(df_fg_market_filtered, col_name)
                        st.metric(f"ยอดขาย {period_name}", f"{value:,.0f}")
                    
                    market_1_month = calculate_period_sales(df_fg_market_filtered, "ยอดขาย 60 วัน", "ยอดขาย 30")
                    market_2_month = calculate_period_sales(df_fg_market_filtered, "ยอดขาย 90 วัน", "ยอดขาย 60 วัน")
                    
                    st.metric("ยอดขาย 1 เดือนก่อน", f"{market_1_month:,.0f}")
                    st.metric("ยอดขาย 2 เดือนก่อน", f"{market_2_month:,.0f}")
                    
                    # Marketing FG sales trend as bar chart
                    if any(col_name in df_fg_market_filtered.columns for col_name, _ in sales_periods):
                        # Create ordered data for proper sequence
                        ordered_periods = [
                            ("ยอดขาย3วัน", "3 วัน"),
                            ("ยอดขาย 7", "7 วัน"), 
                            ("ยอดขาย 15", "15 วัน"),
                            ("ยอดขาย 30", "1 เดือน")
                        ]
                        
                        market_sales_data = []
                        market_labels = []
                        
                        # Add main periods in order
                        for col_name, period_name in ordered_periods:
                            if col_name in df_fg_market_filtered.columns:
                                market_sales_data.append(sales_sum(df_fg_market_filtered, col_name))
                                market_labels.append(period_name)
                        
                        # Add derived periods at the end
                        if market_1_month > 0:
                            market_sales_data.append(market_1_month)
                            market_labels.append("1 เดือนก่อน")
                        if market_2_month > 0:
                            market_sales_data.append(market_2_month)
                            market_labels.append("2 เดือนก่อน")
                        
                        if market_sales_data:
                            st.subheader("📈 แนวโน้มยอดขาย FG ทำตลาด")
                            
                            # Create DataFrame with ordered index
                            market_sales_df = pd.DataFrame({
                                'ช่วงเวลา': [f"{i+1:02d}. {label}" for i, label in enumerate(market_labels)],
                                'ยอดขาย': market_sales_data
                            })
                            
                            # Display as bar chart with ordered index
                            st.bar_chart(market_sales_df.set_index('ช่วงเวลา'))
                            
                            # Display values table for reference (without numbers)
                            st.write("**ค่าตัวเลขยอดขาย:**")
                            for i, (label, value) in enumerate(zip(market_labels, market_sales_data)):
                                st.write(f"**{label}:** {value:,.0f}")
                else:
                    st.info("ไม่พบข้อมูล FG รุ่นทำตลาด")
        
        # Export functionality
        st.divider()
        st.subheader("📁 ส่งออกข้อมูล")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Remove normalized column before export
            export_cols = [col for col in df_sm_all.columns if col != "รหัสรูปแบบ_normalized"]
            df_sm_all[export_cols].to_excel(writer, sheet_name="SM", index=False)
            
            export_cols = [col for col in df_fg_all.columns if col != "รหัสรูปแบบ_normalized"]
            df_fg_all[export_cols].to_excel(writer, sheet_name="FG", index=False)
            
            if len(df_fg_market_filtered) > 0:
                export_cols = [col for col in df_fg_market_filtered.columns if col != "รหัสรูปแบบ_normalized"]
                df_fg_market_filtered[export_cols].to_excel(writer, sheet_name="FG_รุ่นทำตลาด", index=False)
            
            if len(df_sm_use) > 0:
                export_cols = [col for col in df_sm_use.columns if col != "รหัสรูปแบบ_normalized"]
                df_sm_use[export_cols].to_excel(writer, sheet_name="SM_ใช้งาน", index=False)
        
        st.download_button(
            "📁 ดาวน์โหลด Excel",
            output.getvalue(),
            file_name="dashboard_output_fixed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {str(e)}")
        st.error("กรุณาตรวจสอบ:")
        st.write("- ไฟล์ Excel มีครบ 6 ชีต")
        st.write("- ชื่อ Columns ถูกต้อง")
        st.write("- ข้อมูลไม่เสียหาย")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel ที่มีครบ 5 ชีต:")
    st.write("1. **Sheet1** - ข้อมูลหลัก (ช่อง จํานวนที่ใช้ได้ × ราคา)")
    st.write("2. **ราคา** - ข้อมูลราคา")
    st.write("3. **SM** - รายการ SM")
    st.write("4. **ตัดออก** - รายการที่ตัดออก")
    st.write("5. **FG รุ่นทำตลาด** - รายการ FG ทำตลาด")
    st.write("6. **SM ใช้งาน** - รายการ SM ที่ใช้งาน")
    st.write("")
    st.info("📋 **รายการทั้งหมด** = SM + SM WIP + FG + FG WIP (คำนวณจาก ช่อง จํานวนที่ใช้ได้ × ราคา)")