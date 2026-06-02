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
        with st.spinner("Analyzing run-rates and calculating trends by Product Code..."):
            try:
                # --- HELPER TO CLEAN PIVOT TABLE TOTALS ---
                def remove_totals(df, col_name):
                    return df[~df[col_name].astype(str).str.contains('Total', case=False, na=False)]

                # --- 1. PROCESS STOCK ---
                stock_df = pd.read_excel(stock_file, sheet_name='Pivot Table', header=1)
                stock_grouped = stock_df.groupby('Product | Material Code')['Total'].sum().reset_index()
                stock_grouped.rename(columns={'Total': 'Current Stock'}, inplace=True)
                stock_grouped = remove_totals(stock_grouped, 'Product | Material Code')

                # --- 2. PROCESS INCOMING ---
                incoming_df = pd.read_excel(incoming_file, sheet_name='Summary Data')
                incoming_grouped = incoming_df.groupby('Product | Material Code')['Actual Outstanding Quantity'].sum().reset_index()
                incoming_grouped.rename(columns={'Actual Outstanding Quantity': 'Incoming Stock'}, inplace=True)
                incoming_grouped = remove_totals(incoming_grouped, 'Product | Material Code')

                # --- 3. PROCESS SALES & TRENDS ---
                sales_df = pd.read_excel(sales_file, sheet_name='Pivot Table', header=1)
                week_cols = [c for c in sales_df.columns if str(c).isdigit()]
                sales_grouped = sales_df.groupby('Product | Material Code')[week_cols].sum().reset_index()
                sales_grouped = remove_totals(sales_grouped, 'Product | Material Code')
                
                # Get the current week vs the previous weeks average
                current_week = week_cols[-1]
                prev_weeks = week_cols[:-1]
                
                sales_grouped['Current Week Sales'] = sales_grouped[current_week]
                sales_grouped['Prev Weekly Avg'] = sales_grouped[prev_weeks].mean(axis=1)
                sales_grouped['Overall Weekly Avg'] = sales_grouped[week_cols].mean(axis=1)
                
                # Calculate Spike/Drop % and Quantity Change
                sales_grouped['Change %'] = (sales_grouped['Current Week Sales'] - sales_grouped['Prev Weekly Avg']) / sales_grouped['Prev Weekly Avg'].replace(0, pd.NA)
                sales_grouped['Quantity Change'] = sales_grouped['Current Week Sales'] - sales_grouped['Prev Weekly Avg']

                # --- 4. PROCESS PRODUCTION (IF UPLOADED) ---
                if prod_file:
                    prod_df = pd.read_excel(prod_file, sheet_name='Summary Data')
                    prod_df['Total Moved'] = prod_df.get('Quantity In', pd.Series(0)).fillna(0) + prod_df.get('Quantity Out', pd.Series(0)).fillna(0)
                    prod_grouped = prod_df.groupby('Product | Material Code')['Total Moved'].sum().reset_index()
                    prod_grouped['Weekly Prod Usage'] = prod_grouped['Total Moved'] / 4
                    prod_grouped = remove_totals(prod_grouped, 'Product | Material Code')
                else:
                    prod_grouped = pd.DataFrame(columns=['Product | Material Code', 'Weekly Prod Usage'])

                # --- 5. MERGE & CALCULATE ---
                all_codes = pd.DataFrame({'Product | Material Code': pd.concat([
                    stock_grouped['Product | Material Code'], incoming_grouped['Product | Material Code'], 
                    sales_grouped['Product | Material Code'], prod_grouped['Product | Material Code']
                ]).unique()})
                # Final safety check to ensure no blank/total rows snuck through
                all_codes = all_codes.dropna()

                df = all_codes.merge(stock_grouped, on='Product | Material Code', how='left')\
                             .merge(incoming_grouped, on='Product | Material Code', how='left')\
                             .merge(sales_grouped[['Product | Material Code', 'Overall Weekly Avg', 'Current Week Sales', 'Prev Weekly Avg', 'Change %', 'Quantity Change']], on='Product | Material Code', how='left')\
                             .merge(prod_grouped[['Product | Material Code', 'Weekly Prod Usage']], on='Product | Material Code', how='left')

                df.fillna({'Current Stock': 0, 'Incoming Stock': 0, 'Overall Weekly Avg': 0, 'Weekly Prod Usage': 0, 'Current Week Sales': 0, 'Prev Weekly Avg': 0, 'Quantity Change': 0}, inplace=True)

                df['Total Weekly Demand'] = df['Overall Weekly Avg'] + df['Weekly Prod Usage']
                df['Total Expected Stock'] = df['Current Stock'] + df['Incoming Stock']
                
                def calc_wos(row):
                    if row['Total Weekly Demand'] <= 0: return 999 if row['Total Expected Stock'] > 0 else 0
                    return row['Total Expected Stock'] / row['Total Weekly Demand']
                df['WOS'] = df.apply(calc_wos, axis=1)

                # Format percentages beautifully for display
                df['Change % Display'] = (df['Change %'] * 100).fillna(0).round(1).astype(str) + "%"

                # --- 6. CREATE THE 6 ACTIONABLE LISTS WITH CUSTOM LOGIC ---
                
                # 1. Dead Stock (Sorted by highest inventory, now showing total expected stock)
                cat1 = df[(df['Total Weekly Demand'] == 0) & (df['Current Stock'] > 0)].sort_values('Current Stock', ascending=False)
                
                # 2. Slow Movers (Sorted by highest WOS instead of current stock)
                cat2 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] > 10) & (df['WOS'] != 999)].sort_values('WOS', ascending=False)
                
                # 3. Understock Risk (Sorted by LOWEST WOS to highlight most urgent, then highest demand)
                cat3 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] < 4)].sort_values(['WOS', 'Total Weekly Demand'], ascending=[True, False])
                
                # 4. Reorder Needed (Sorted by LOWEST WOS to highlight closest to running out, Total row removed)
                cat4 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] >= 4) & (df['WOS'] <= 10) & (df['Incoming Stock'] == 0)].sort_values('WOS', ascending=True)
                
                # 5. Spikes (Sorted by highest Quantity Change)
                cat5 = df[df['Change %'] > 0.3].sort_values('Quantity Change', ascending=False)
                
                # 6. Drops (Sorted by lowest Quantity Change, pushing Current Stock = 0 to the very bottom)
                cat6 = df[df['Change %'] < -0.3].copy()
                cat6['Is_Out_Of_Stock'] = cat6['Current Stock'] <= 0
                # Sort first by Is_Out_Of_Stock (False comes before True), then by Quantity Change (most negative first)
                cat6 = cat6.sort_values(['Is_Out_Of_Stock', 'Quantity Change'], ascending=[True, True])

                # --- 7. EXPORT TO EXCEL FEATURE ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    cat1[['Product | Material Code', 'Current Stock', 'Incoming Stock', 'Total Expected Stock']].to_excel(writer, sheet_name="1. Dead Stock", index=False)
                    cat2[['Product | Material Code', 'Current Stock', 'Total Weekly Demand', 'WOS']].to_excel(writer, sheet_name="2. Slow Movers", index=False)
                    cat3[['Product | Material Code', 'WOS', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Total Weekly Demand']].to_excel(writer, sheet_name="3. Understock Risk", index=False)
                    cat4[['Product | Material Code', 'WOS', 'Current Stock', 'Total Weekly Demand']].to_excel(writer, sheet_name="4. Reorder Needed", index=False)
                    cat5[['Product | Material Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']].to_excel(writer, sheet_name="5. Sales Spikes", index=False)
                    cat6[['Product | Material Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']].to_excel(writer, sheet_name="6. Sales Drops", index=False)
                buffer.seek(0)

                # --- 8. DISPLAY ON SCREEN ---
                st.success("Actionable items generated successfully!")
                st.download_button(label="📥 Download Action Items to Excel", data=buffer, file_name="Inventory_Action_Items_By_Product.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                t1, t2, t3, t4, t5, t6 = st.tabs(["1. Dead Stock", "2. Overstock/Slow", "3. Understock", "4. Reorder Watch", "5. Spikes (+30%)", "6. Drops (-30%)"])
                
                with t1: st.dataframe(cat1[['Product | Material Code', 'Current Stock', 'Incoming Stock', 'Total Expected Stock']], use_container_width=True)
                with t2: st.dataframe(cat2[['Product | Material Code', 'Current Stock', 'Total Weekly Demand', 'WOS']], use_container_width=True)
                with t3: st.dataframe(cat3[['Product | Material Code', 'WOS', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Total Weekly Demand']], use_container_width=True)
                with t4: st.dataframe(cat4[['Product | Material Code', 'WOS', 'Current Stock', 'Total Weekly Demand']], use_container_width=True)
                with t5: st.dataframe(cat5[['Product | Material Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']], use_container_width=True)
                with t6: st.dataframe(cat6[['Product | Material Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']], use_container_width=True)

            except Exception as e:
                st.error(f"Error processing files: {e}")
