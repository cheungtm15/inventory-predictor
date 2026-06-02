import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Actionable Inventory", page_icon="🎯", layout="wide")
st.title("🎯 Actionable Inventory Engine")
st.write("Upload your Excel files to generate your daily 6-step actionable inventory to-do list.")

col1, col2, col3, col4 = st.columns(4)
with col1:
    stock_file = st.file_uploader("1. Stock Level (.xlsx)", type=['xlsx'])
with col2:
    sales_file = st.file_uploader("2. Sales by SKU (.xlsx)", type=['xlsx'])
with col3:
    incoming_file = st.file_uploader("3. Incoming Shipments (.xlsx)", type=['xlsx'])
with col4:
    prod_file = st.file_uploader("4. Production Items (.xlsx)", type=['xlsx'])

if stock_file and sales_file and incoming_file:
    if st.button("Generate Action Items", type="primary"):
        with st.spinner("Analyzing run-rates and calculating trends..."):
            try:
                # --- 1. PROCESS STOCK ---
                stock_df = pd.read_excel(stock_file, sheet_name='Pivot Table', header=1)
                stock_grouped = stock_df.groupby('Product | Material Base SKU')['Total'].sum().reset_index()
                stock_grouped.rename(columns={'Total': 'Current Stock'}, inplace=True)

                # --- 2. PROCESS INCOMING ---
                incoming_df = pd.read_excel(incoming_file, sheet_name='Summary Data')
                incoming_grouped = incoming_df.groupby('Product | Material Base SKU')['Actual Outstanding Quantity'].sum().reset_index()
                incoming_grouped.rename(columns={'Actual Outstanding Quantity': 'Incoming Stock'}, inplace=True)

                # --- 3. PROCESS SALES & TRENDS ---
                sales_df = pd.read_excel(sales_file, sheet_name='Pivot Table', header=1)
                week_cols = [c for c in sales_df.columns if str(c).isdigit()]
                sales_grouped = sales_df.groupby('Product | Material Base SKU')[week_cols].sum().reset_index()
                
                # Get the current week vs the previous weeks average
                current_week = week_cols[-1]
                prev_weeks = week_cols[:-1]
                
                sales_grouped['Current Week Sales'] = sales_grouped[current_week]
                sales_grouped['Prev Weekly Avg'] = sales_grouped[prev_weeks].mean(axis=1)
                sales_grouped['Overall Weekly Avg'] = sales_grouped[week_cols].mean(axis=1)
                
                # Calculate Spike/Drop %
                sales_grouped['Change %'] = (sales_grouped['Current Week Sales'] - sales_grouped['Prev Weekly Avg']) / sales_grouped['Prev Weekly Avg'].replace(0, pd.NA)

                # --- 4. PROCESS PRODUCTION (IF UPLOADED) ---
                if prod_file:
                    prod_df = pd.read_excel(prod_file, sheet_name='Summary Data')
                    prod_df['Total Moved'] = prod_df.get('Quantity In', pd.Series(0)).fillna(0) + prod_df.get('Quantity Out', pd.Series(0)).fillna(0)
                    prod_grouped = prod_df.groupby('Product | Material Base SKU')['Total Moved'].sum().reset_index()
                    prod_grouped['Weekly Prod Usage'] = prod_grouped['Total Moved'] / 4
                else:
                    prod_grouped = pd.DataFrame(columns=['Product | Material Base SKU', 'Weekly Prod Usage'])

                # --- 5. MERGE & CALCULATE ---
                all_skus = pd.DataFrame({'Product | Material Base SKU': pd.concat([
                    stock_grouped['Product | Material Base SKU'], incoming_grouped['Product | Material Base SKU'], 
                    sales_grouped['Product | Material Base SKU'], prod_grouped['Product | Material Base SKU']
                ]).unique()})

                df = all_skus.merge(stock_grouped, on='Product | Material Base SKU', how='left')\
                             .merge(incoming_grouped, on='Product | Material Base SKU', how='left')\
                             .merge(sales_grouped[['Product | Material Base SKU', 'Overall Weekly Avg', 'Current Week Sales', 'Prev Weekly Avg', 'Change %']], on='Product | Material Base SKU', how='left')\
                             .merge(prod_grouped[['Product | Material Base SKU', 'Weekly Prod Usage']], on='Product | Material Base SKU', how='left')

                df.fillna({'Current Stock': 0, 'Incoming Stock': 0, 'Overall Weekly Avg': 0, 'Weekly Prod Usage': 0, 'Current Week Sales': 0, 'Prev Weekly Avg': 0}, inplace=True)

                df['Total Weekly Demand'] = df['Overall Weekly Avg'] + df['Weekly Prod Usage']
                df['Expected Stock'] = df['Current Stock'] + df['Incoming Stock']
                
                def calc_wos(row):
                    if row['Total Weekly Demand'] <= 0: return 999 if row['Expected Stock'] > 0 else 0
                    return row['Expected Stock'] / row['Total Weekly Demand']
                df['WOS'] = df.apply(calc_wos, axis=1)

                # Helper to format percentages beautifully
                df['Change % Display'] = (df['Change %'] * 100).fillna(0).round(1).astype(str) + "%"

                # --- 6. CREATE THE 6 ACTIONABLE LISTS ---
                
                # 1. No sales, but with inventory (Sorted by highest inventory)
                cat1 = df[(df['Total Weekly Demand'] == 0) & (df['Current Stock'] > 0)].sort_values('Current Stock', ascending=False)
                
                # 2. Low sales, with inventory (WOS > 10) (Sorted by highest inventory)
                cat2 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] > 10) & (df['WOS'] != 999)].sort_values('Current Stock', ascending=False)
                
                # 3. High sales, low inventory (WOS < 4) (Sorted by highest demand, showing incoming)
                cat3 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] < 4)].sort_values('Total Weekly Demand', ascending=False)
                
                # 4. High sales, normal inventory (WOS 4-10) BUT no incoming shipment (Sorted by highest demand)
                cat4 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] >= 4) & (df['WOS'] <= 10) & (df['Incoming Stock'] == 0)].sort_values('Total Weekly Demand', ascending=False)
                
                # 5. Sudden Increase (> 30%)
                cat5 = df[df['Change %'] > 0.3].sort_values('Change %', ascending=False)
                
                # 6. Sudden Drop (< -30%)
                cat6 = df[df['Change %'] < -0.3].sort_values('Change %', ascending=True)

                # --- 7. EXPORT TO EXCEL FEATURE ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    cat1[['Product | Material Base SKU', 'Current Stock', 'Incoming Stock']].to_excel(writer, sheet_name="1. Dead Stock", index=False)
                    cat2[['Product | Material Base SKU', 'Current Stock', 'Total Weekly Demand', 'WOS']].to_excel(writer, sheet_name="2. Slow Movers", index=False)
                    cat3[['Product | Material Base SKU', 'Total Weekly Demand', 'Current Stock', 'Incoming Stock', 'WOS']].to_excel(writer, sheet_name="3. Understock Risk", index=False)
                    cat4[['Product | Material Base SKU', 'Total Weekly Demand', 'Current Stock', 'WOS']].to_excel(writer, sheet_name="4. Reorder Needed", index=False)
                    cat5[['Product | Material Base SKU', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']].to_excel(writer, sheet_name="5. Sales Spikes", index=False)
                    cat6[['Product | Material Base SKU', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']].to_excel(writer, sheet_name="6. Sales Drops", index=False)
                buffer.seek(0)

                # --- 8. DISPLAY ON SCREEN ---
                st.success("Actionable items generated successfully!")
                st.download_button(label="📥 Download Action Items to Excel", data=buffer, file_name="Inventory_Action_Items.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                t1, t2, t3, t4, t5, t6 = st.tabs(["1. Dead Stock", "2. Overstock/Slow", "3. Understock", "4. Reorder Watch", "5. Spikes (+30%)", "6. Drops (-30%)"])
                
                display_cols_standard = ['Product | Material Base SKU', 'Current Stock', 'Incoming Stock', 'Total Weekly Demand', 'WOS']
                display_cols_trend = ['Product | Material Base SKU', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']

                with t1: st.dataframe(cat1[display_cols_standard], use_container_width=True)
                with t2: st.dataframe(cat2[display_cols_standard], use_container_width=True)
                with t3: st.dataframe(cat3[display_cols_standard], use_container_width=True)
                with t4: st.dataframe(cat4[display_cols_standard], use_container_width=True)
                with t5: st.dataframe(cat5[display_cols_trend], use_container_width=True)
                with t6: st.dataframe(cat6[display_cols_trend], use_container_width=True)

            except Exception as e:
                st.error(f"Error processing files: {e}")
