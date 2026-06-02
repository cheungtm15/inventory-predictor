import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Actionable Inventory", page_icon="🎯", layout="wide")
st.title("🎯 Actionable Inventory Engine")
st.write("Upload your Excel files to generate your daily actionable to-do list and Master Overviews.")

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
        with st.spinner("Analyzing run-rates, formatting Excel sheets, and building Master Lists..."):
            try:
                # --- HELPER TO CLEAN PIVOT TABLE TOTALS ---
                def remove_totals(df, col_name):
                    return df[~df[col_name].astype(str).str.contains('Total', case=False, na=False)]

                # --- 1. PROCESS STOCK (Extract Base SKU & Brand) ---
                stock_df = pd.read_excel(stock_file, sheet_name='Pivot Table', header=1)
                stock_df = remove_totals(stock_df, 'Product | Material Code')
                
                stock_grouped = stock_df.groupby('Product | Material Code')['Total'].sum().reset_index()
                stock_grouped.rename(columns={'Total': 'Current Stock'}, inplace=True)
                
                # Extract both Brand AND Base SKU for Sheet 7 and 8
                info_mapping = stock_df.groupby('Product | Material Code')[['Product | Material Base SKU', 'Product | Material Brand']].first().reset_index()
                stock_grouped = stock_grouped.merge(info_mapping, on='Product | Material Code', how='left')
                stock_grouped['In_Stock_Report'] = True

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
                
                current_week = week_cols[-1]
                prev_weeks = week_cols[:-1]
                
                sales_grouped['Current Week Sales'] = sales_grouped[current_week]
                sales_grouped['Prev Weekly Avg'] = sales_grouped[prev_weeks].mean(axis=1)
                sales_grouped['Overall Weekly Avg'] = sales_grouped[week_cols].mean(axis=1)
                
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

                # --- 5. MERGE & CALCULATE MULTI-LAYER DATA ---
                all_codes = pd.DataFrame({'Product | Material Code': pd.concat([
                    stock_grouped['Product | Material Code'], incoming_grouped['Product | Material Code'], 
                    sales_grouped['Product | Material Code'], prod_grouped['Product | Material Code']
                ]).unique()})
                all_codes = all_codes.dropna()

                df = all_codes.merge(stock_grouped, on='Product | Material Code', how='left')\
                             .merge(incoming_grouped, on='Product | Material Code', how='left')\
                             .merge(sales_grouped[['Product | Material Code', 'Overall Weekly Avg', 'Current Week Sales', 'Prev Weekly Avg', 'Change %', 'Quantity Change']], on='Product | Material Code', how='left')\
                             .merge(prod_grouped[['Product | Material Code', 'Weekly Prod Usage']], on='Product | Material Code', how='left')

                df.fillna({
                    'Current Stock': 0, 'Incoming Stock': 0, 'Overall Weekly Avg': 0, 
                    'Weekly Prod Usage': 0, 'Current Week Sales': 0, 'Prev Weekly Avg': 0, 
                    'Quantity Change': 0, 'In_Stock_Report': False, 'Product | Material Brand': 'Unknown', 'Product | Material Base SKU': 'Unknown'
                }, inplace=True)

                df['Total Weekly Demand'] = df['Overall Weekly Avg'] + df['Weekly Prod Usage']
                df['Total Expected Stock'] = df['Current Stock'] + df['Incoming Stock']
                
                def calc_wos(row):
                    if row['Total Weekly Demand'] <= 0: return 999 if row['Total Expected Stock'] > 0 else 0
                    if row['Total Expected Stock'] <= 0: return 0.0 
                    return row['Total Expected Stock'] / row['Total Weekly Demand']
                df['WOS'] = df.apply(calc_wos, axis=1)
                df['Change % Display'] = (df['Change %'] * 100).fillna(0).round(1).astype(str) + "%"

                # --- 6. CREATE THE 6 ACTIONABLE TABS ---
                cat1 = df[(df['Total Weekly Demand'] == 0) & (df['Current Stock'] > 0)].sort_values('Current Stock', ascending=False).copy()
                cat1['Action Recommended'] = cat1['Incoming Stock'].apply(lambda x: "🚨 Review PO" if x > 0 else "Hold / Discount")
                
                cat2 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] > 10) & (df['WOS'] != 999)].sort_values('WOS', ascending=False).copy()
                cat2['Action Recommended'] = cat2['Incoming Stock'].apply(lambda x: "🚨 Review PO" if x > 0 else "Monitor")
                
                cat3 = df[(df['Total Weekly Demand'] > 0) & (df['Total Expected Stock'] / df['Total Weekly Demand'] < 4) & (df['In_Stock_Report'] == True)].copy()
                def get_risk_level(row):
                    if row['Total Expected Stock'] <= 0: return "1. 🚨 OUT OF STOCK / NEGATIVE"
                    elif row['WOS'] <= 2: return "2. 🔴 CRITICAL (< 2 Weeks)"
                    else: return "3. 🟡 LOW STOCK (2-4 Weeks)"
                cat3['Risk Level'] = cat3.apply(get_risk_level, axis=1)
                cat3 = cat3.sort_values(['Risk Level', 'Total Weekly Demand'], ascending=[True, False])
                
                cat4 = df[(df['Total Weekly Demand'] > 0) & (df['WOS'] >= 4) & (df['WOS'] <= 10) & (df['Incoming Stock'] == 0) & (df['In_Stock_Report'] == True)]\
                        .sort_values(['Product | Material Brand', 'WOS'], ascending=[True, True])
                
                cat5 = df[df['Change %'] > 0.3].sort_values('Quantity Change', ascending=False)
                
                cat6 = df[df['Change %'] < -0.3].copy()
                cat6['Is_Out_Of_Stock'] = cat6['Current Stock'] <= 0
                cat6 = cat6.sort_values(['Is_Out_Of_Stock', 'Quantity Change'], ascending=[True, True])

                # --- 7. CREATE MASTER SHEET 7 (PRODUCT LEVEL) ---
                def get_inv_status(row):
                    if row['Total Weekly Demand'] == 0 and row['Current Stock'] > 0: return "Dead Stock"
                    elif row['Total Weekly Demand'] > 0 and row['Total Expected Stock']/row['Total Weekly Demand'] < 4: return "Understock Risk"
                    elif row['Total Weekly Demand'] > 0 and row['WOS'] > 10 and row['WOS'] != 999: return "Slow Mover"
                    elif row['Total Weekly Demand'] > 0 and 4 <= row['WOS'] <= 10 and row['Incoming Stock'] == 0: return "Reorder Needed"
                    elif row['Total Weekly Demand'] > 0 and 4 <= row['WOS'] <= 10 and row['Incoming Stock'] > 0: return "Healthy (Incoming Planned)"
                    elif row['Total Weekly Demand'] == 0 and row['Total Expected Stock'] <= 0: return "Out of Stock & No Demand"
                    else: return "Unknown"
                    
                def get_trend_status(row):
                    if pd.isna(row['Change %']): return "Stable"
                    elif row['Change %'] > 0.3: return "Spike (>30%)"
                    elif row['Change %'] < -0.3: return "Drop (<-30%)"
                    else: return "Stable"

                sheet7 = df[df['In_Stock_Report'] == True].copy()
                sheet7.rename(columns={'Overall Weekly Avg': 'Weekly Avg Sales'}, inplace=True)
                sheet7['Inventory Status'] = sheet7.apply(get_inv_status, axis=1)
                sheet7['Sales Trend'] = sheet7.apply(get_trend_status, axis=1)
                sheet7 = sheet7.sort_values(['Product | Material Brand', 'Product | Material Code'])
                
                sheet7_cols = ['Product | Material Brand', 'Product | Material Base SKU', 'Product | Material Code', 'Inventory Status', 'Sales Trend', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Weekly Avg Sales', 'Weekly Prod Usage', 'Total Weekly Demand', 'WOS', 'Quantity Change', 'Change % Display']

                # --- 8. CREATE MASTER SHEET 8 (BASE SKU LEVEL) ---
                sheet8_raw = df[df['In_Stock_Report'] == True].groupby('Product | Material Base SKU').agg({
                    'Current Stock': 'sum', 'Incoming Stock': 'sum', 'Overall Weekly Avg': 'sum',
                    'Weekly Prod Usage': 'sum', 'Current Week Sales': 'sum', 'Prev Weekly Avg': 'sum', 'Quantity Change': 'sum'
                }).reset_index()
                
                sheet8_raw.rename(columns={'Overall Weekly Avg': 'Weekly Avg Sales'}, inplace=True)
                sheet8_raw['Total Weekly Demand'] = sheet8_raw['Weekly Avg Sales'] + sheet8_raw['Weekly Prod Usage']
                sheet8_raw['Total Expected Stock'] = sheet8_raw['Current Stock'] + sheet8_raw['Incoming Stock']
                
                # Recalculate WOS and Change % as an aggregate Whole SKU
                sheet8_raw['WOS'] = sheet8_raw.apply(calc_wos, axis=1)
                sheet8_raw['Change %'] = (sheet8_raw['Current Week Sales'] - sheet8_raw['Prev Weekly Avg']) / sheet8_raw['Prev Weekly Avg'].replace(0, pd.NA)
                sheet8_raw['Change % Display'] = (sheet8_raw['Change %'] * 100).fillna(0).round(1).astype(str) + "%"
                
                sheet8_raw['Inventory Status'] = sheet8_raw.apply(get_inv_status, axis=1)
                sheet8_raw['Sales Trend'] = sheet8_raw.apply(get_trend_status, axis=1)
                
                sheet8 = sheet8_raw.sort_values('Product | Material Base SKU')
                sheet8_cols = ['Product | Material Base SKU', 'Inventory Status', 'Sales Trend', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Weekly Avg Sales', 'Weekly Prod Usage', 'Total Weekly Demand', 'WOS', 'Quantity Change', 'Change % Display']

                # --- 9. EXPORT TO EXCEL ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    cat1[['Product | Material Code', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Action Recommended']].to_excel(writer, sheet_name="1. Dead Stock", index=False)
                    cat2[['Product | Material Code', 'Current Stock', 'Total Weekly Demand', 'WOS', 'Action Recommended']].to_excel(writer, sheet_name="2. Slow Movers", index=False)
                    cat3[['Risk Level', 'Product | Material Brand', 'Product | Material Code', 'WOS', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Total Weekly Demand']].to_excel(writer, sheet_name="3. Understock Risk", index=False)
                    cat4[['Product | Material Brand', 'Product | Material Code', 'WOS', 'Current Stock', 'Total Weekly Demand']].to_excel(writer, sheet_name="4. Reorder Needed", index=False)
                    cat5[['Product | Material Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']].to_excel(writer, sheet_name="5. Sales Spikes", index=False)
                    cat6[['Product | Material Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']].to_excel(writer, sheet_name="6. Sales Drops", index=False)
                    
                    # Write Master Sheets
                    sheet7[sheet7_cols].to_excel(writer, sheet_name="7. Master (Product Code)", index=False)
                    sheet8[sheet8_cols].to_excel(writer, sheet_name="8. Master (Base SKU)", index=False)
                    
                    # Auto-width formatting for ALL 8 sheets
                    for sheetname, worksheet in writer.sheets.items():
                        worksheet.freeze_panes = 'A2'
                        for col in worksheet.columns:
                            max_length = 0
                            column_letter = col[0].column_letter
                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                                except: pass
                            worksheet.column_dimensions[column_letter].width = (max_length + 2)

                buffer.seek(0)

                # --- 10. DISPLAY ON SCREEN ---
                st.success("Actionable items and Master Lists generated successfully!")
                st.download_button(label="📥 Download Formatted Excel Report", data=buffer, file_name="Inventory_Action_Items.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                t1, t2, t3, t4, t5, t6, t7, t8 = st.tabs(["1. Dead", "2. Slow", "3. Understock", "4. Reorder", "5. Spikes", "6. Drops", "7. Master (Product)", "8. Master (SKU)"])
                
                with t1: st.dataframe(cat1[['Product | Material Code', 'Current Stock', 'Incoming Stock', 'Action Recommended']], use_container_width=True)
                with t2: st.dataframe(cat2[['Product | Material Code', 'Current Stock', 'WOS', 'Action Recommended']], use_container_width=True)
                with t3: st.dataframe(cat3[['Risk Level', 'Product | Material Brand', 'Product | Material Code', 'WOS', 'Total Weekly Demand']], use_container_width=True)
                with t4: st.dataframe(cat4[['Product | Material Brand', 'Product | Material Code', 'WOS', 'Total Weekly Demand']], use_container_width=True)
                with t5: st.dataframe(cat5[['Product | Material Code', 'Quantity Change', 'Current Stock']], use_container_width=True)
                with t6: st.dataframe(cat6[['Product | Material Code', 'Quantity Change', 'Current Stock']], use_container_width=True)
                with t7: st.dataframe(sheet7[sheet7_cols], use_container_width=True)
                with t8: st.dataframe(sheet8[sheet8_cols], use_container_width=True)

            except Exception as e:
                st.error(f"Error processing files: {e}")
