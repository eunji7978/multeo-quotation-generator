import streamlit as st
import json
import os
import pandas as pd
from scripts.generate_excel import create_quotation
from datetime import datetime

# --- Constants ---
PRODUCTS_FILE = "assets/products.json"

# --- Helper Functions ---
def load_products():
    if not os.path.exists(PRODUCTS_FILE):
        return []
    with open(PRODUCTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_products(products):
    with open(PRODUCTS_FILE, "w", encoding="utf-8") as f:
        json.dump(products, f, ensure_ascii=False, indent=2)

# --- App Layout ---
st.set_page_config(page_title="Multeo Quotation Generator", layout="wide")
st.title("Multeo 견적서 생성기")

tab1, tab2 = st.tabs(["💰 견적서 작성", "⚙️ 품목 관리"])

# --- Tab 1: Quotation Maker ---
with tab1:
    st.header("견적서 작성")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        recipient_name = st.text_input("받는 사람 (업체명/성명)", value="레퍼토리 성수")
    
    products = load_products()
    product_names = [p['name'] for p in products]
    
    st.subheader("품목 선택")
    
    # Session state to keep track of selected items
    if 'selected_items' not in st.session_state:
        st.session_state.selected_items = []

    # Add product interface
    with st.expander("품목 추가하기", expanded=True):
        selected_product_name = st.selectbox("품목을 선택하세요", options=[""] + product_names)
        qty_input = st.number_input("수량", min_value=1, value=1)
        
        if st.button("추가"):
            if selected_product_name:
                # Find product details
                prod = next((p for p in products if p['name'] == selected_product_name), None)
                if prod:
                    st.session_state.selected_items.append({
                        "name": prod['name'],
                        "unit_price": prod['price'],
                        "quantity": qty_input
                    })
                    st.success(f"{selected_product_name} 추가됨")
                else:
                    st.error("품목을 찾을 수 없습니다.")
            else:
                st.warning("품목을 선택해주세요.")

    # Show selected items list
    if st.session_state.selected_items:
        st.subheader("견적 품목 리스트")
        
        # Convert to DataFrame for display (allow editing quantity roughly? No, Streamlit data_editor is better)
        df_items = pd.DataFrame(st.session_state.selected_items)
        
        # Calculate supply price for display
        df_items['supply_price'] = (df_items['unit_price'] * 0.6).astype(int)
        df_items['total'] = df_items['supply_price'] * df_items['quantity']
        
        edited_df = st.data_editor(df_items, num_rows="dynamic", key="editor")
        
        # Update session state from editor (handle deletions/edits)
        # Note: data_editor returns the new dataframe.
        # We need to sync back to session_state for logic usage.
        
        total_estimate = edited_df['total'].sum()
        st.metric("총 견적 금액 (공급가액 합계)", f"{total_estimate:,} 원")
        
        if st.button("견적서 엑셀 생성"):
            base_filename = f"견적서_{recipient_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_path = os.path.join("downloads", base_filename)
            os.makedirs("downloads", exist_ok=True)
            
            # Prepare items list from edited dataframe
            final_items = []
            for index, row in edited_df.iterrows():
                final_items.append({
                    "name": row['name'],
                    "unit_price": int(row['unit_price']),
                    "quantity": int(row['quantity'])
                })
            
            create_quotation(recipient_name, final_items, output_path)
            
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=f,
                    file_name=base_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("품목을 추가해주세요.")

# --- Tab 2: Product Manager ---
with tab2:
    st.header("품목 및 단가 관리")
    
    current_products = load_products()
    df_products = pd.DataFrame(current_products)
    
    st.write("아래 표에서 품목명과 가격을 직접 수정할 수 있습니다.")
    edited_products_df = st.data_editor(df_products, num_rows="dynamic")
    
    if st.button("변경사항 저장"):
        # Convert back to list of dicts
        updated_products = edited_products_df.to_dict(orient="records")
        save_products(updated_products)
        st.success("저장되었습니다!")
        st.rerun()
