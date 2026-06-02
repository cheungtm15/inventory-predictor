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
        with st.spinner("Applying heatmaps, calculating weeks of supply, and cleaning data..."):
            try:
                # --- HELPER TO CLEAN PIVOT TABLE TOTALS ---
                def remove_totals(df, col_name):
                    return df[~df[col_name].astype(str).str.contains('Total', case=False, na=False)]

                # Global rename map to clean up headers
                rename_map = {
                    'Product | Material Code': 'Code',
                    'Product | Material Base SKU': 'Base SKU',
                    'Product | Material Brand': 'Brand',
                    'Product | Material Series Description': 'Series'
                }

                # --- 1. PROCESS STOCK ---
                stock_df = pd.read_excel(stock_file, sheet_name='Pivot Table', header=1)
                stock_df = stock_df.rename(columns=rename_map)
                stock_df = remove_totals(stock_df, 'Code')
                
                stock_grouped = stock_df.groupby('Code')['Total'].sum().reset_index()
                stock_grouped.rename(columns={'Total': 'Current Stock'}, inplace=True)
                
                info_mapping = stock_df.groupby('Code')[['Base SKU', 'Brand', 'Series']].first().reset_index()
                stock_grouped = stock_grouped.merge(info_mapping, on='Code', how='left')
                stock_grouped['In_Stock_Report'] = True

                # --- 2. PROCESS INCOMING ---
                incoming_df = pd.read_excel(incoming_file, sheet_name='Summary Data')
                incoming_df = incoming_df.rename(columns=rename_map)
                incoming_grouped = incoming_df.groupby('Code')['Actual Outstanding Quantity'].sum().reset_index()
                incoming_grouped.rename(columns={'Actual Outstanding Quantity': 'Incoming Stock'}, inplace=True)
                incoming_grouped = remove_totals(incoming_grouped, 'Code')

                # --- 3. PROCESS SALES & TRENDS ---
                sales_df = pd.read_excel(sales_file, sheet_name='Pivot Table', header=1)
                sales_df = sales_df.rename(columns=rename_map)
                week_cols = [c for c in sales_df.columns if str(c).isdigit()]
                sales_grouped = sales_df.groupby('Code')[week_cols].sum().reset_index()
                sales_grouped = remove_totals(sales_grouped, 'Code')
                
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
                    prod_df = prod_df.rename(columns=rename_map)
                    prod_df['Total Moved'] = prod_df.get('Quantity In', pd.Series(0)).fillna(0) + prod_df.get('Quantity Out', pd.Series(0)).fillna(0)
                    prod_grouped = prod_df.groupby('Code')['Total Moved'].sum().reset_index()
                    prod_grouped['Weekly Prod Usage'] = prod_grouped['Total Moved'] / 4
                    prod_grouped = remove_totals(prod_grouped, 'Code')
                else:
                    prod_grouped = pd.DataFrame(columns=['Code', 'Weekly Prod Usage'])

                # --- 5. MERGE & CALCULATE MULTI-LAYER DATA ---
                all_codes = pd.DataFrame({'Code': pd.concat([
                    stock_grouped['Code'], incoming_grouped['Code'], 
                    sales_grouped['Code'], prod_grouped['Code']
                ]).unique()})
                all_codes = all_codes.dropna()

                df = all_codes.merge(stock_grouped, on='Code', how='left')\
                             .merge(incoming_grouped, on='Code', how='left')\
                             .merge(sales_grouped[['Code', 'Overall Weekly Avg', 'Current Week Sales', 'Prev Weekly Avg', 'Change %', 'Quantity Change']], on='Code', how='left')\
                             .merge(prod_grouped[['Code', 'Weekly Prod Usage']], on='Code', how='left')

                df.fillna({
                    'Current Stock': 0, 'Incoming Stock': 0, 'Overall Weekly Avg': 0, 
                    'Weekly Prod Usage': 0, 'Current Week Sales': 0, 'Prev Weekly Avg': 0, 
                    'Quantity Change': 0, 'In_Stock_Report': False, 'Brand': 'Unknown', 
                    'Base SKU': 'Unknown', 'Series': 'Unknown'
                }, inplace=True)

                # Cleaned up KPI Names
                df['Avg. Weekly Demand'] = df['Overall Weekly Avg'] + df['Weekly Prod Usage']
                df['Total Expected Stock'] = df['Current Stock'] + df['Incoming Stock']
                
                def calc_wos(row):
                    if row['Avg. Weekly Demand'] <= 0: return 999 if row['Total Expected Stock'] > 0 else 0
                    if row['Total Expected Stock'] <= 0: return 0.0 
                    return row['Total Expected Stock'] / row['Avg. Weekly Demand']
                
                df['Weeks of Supply'] = df.apply(calc_wos, axis=1)
                df['Change % Display'] = (df['Change %'] * 100).fillna(0).round(1).astype(str) + "%"

                # --- 6. CREATE THE 6 ACTIONABLE TABS ---
                cat1 = df[(df['Avg. Weekly Demand'] == 0) & (df['Current Stock'] > 0)].copy()
                cat1['Action Recommended'] = cat1['Incoming Stock'].apply(lambda x: "🚨 Review PO" if x > 0 else "Hold / Discount")
                cat1 = cat1.sort_values(['Series', 'Current Stock'], ascending=[True, False])
                
                cat2 = df[(df['Avg. Weekly Demand'] > 0) & (df['Weeks of Supply'] > 10) & (df['Weeks of Supply'] != 999)].copy()
                cat2['Action Recommended'] = cat2['Incoming Stock'].apply(lambda x: "🚨 Review PO" if x > 0 else "Monitor")
                cat2 = cat2.sort_values(['Series', 'Current Stock'], ascending=[True, False])
                
                cat3 = df[(df['Avg. Weekly Demand'] > 0) & (df['Total Expected Stock'] / df['Avg. Weekly Demand'] < 4) & (df['In_Stock_Report'] == True)].copy()
                def get_risk_level(row):
                    if row['Total Expected Stock'] <= 0: return "1. 🚨 OUT OF STOCK / NEGATIVE"
                    elif row['Weeks of Supply'] <= 2: return "2. 🔴 CRITICAL (< 2 Weeks)"
                    else: return "3. 🟡 LOW STOCK (2-4 Weeks)"
                cat3['Risk Level'] = cat3.apply(get_risk_level, axis=1)
                cat3 = cat3.sort_values(['Risk Level', 'Avg. Weekly Demand'], ascending=[True, False])
                
                cat4 = df[(df['Avg. Weekly Demand'] > 0) & (df['Weeks of Supply'] >= 4) & (df['Weeks of Supply'] <= 10) & (df['Incoming Stock'] == 0) & (df['In_Stock_Report'] == True)]\
                        .sort_values(['Brand', 'Weeks of Supply'], ascending=[True, True])
                
                cat5 = df[df['Change %'] > 0.3].sort_values('Quantity Change', ascending=False)
                
                cat6 = df[df['Change %'] < -0.3].copy()
                cat6['Is_Out_Of_Stock'] = cat6['Current Stock'] <= 0
                cat6 = cat6.sort_values(['Is_Out_Of_Stock', 'Quantity Change'], ascending=[True, True])

                # --- 7. CREATE MASTER SHEETS (Filtered for noise) ---
                def get_inv_status(row):
                    if row['Avg. Weekly Demand'] == 0 and row['Current Stock'] > 0: return "Dead Stock"
                    elif row['Avg. Weekly Demand'] > 0 and row['Total Expected Stock']/row['Avg. Weekly Demand'] < 4: return "Understock Risk"
                    elif row['Avg. Weekly Demand'] > 0 and row['Weeks of Supply'] > 10 and row['Weeks of Supply'] != 999: return "Slow Mover"
                    elif row['Avg. Weekly Demand'] > 0 and 4 <= row['Weeks of Supply'] <= 10 and row['Incoming Stock'] == 0: return "Reorder Needed"
                    elif row['Avg. Weekly Demand'] > 0 and 4 <= row['Weeks of Supply'] <= 10 and row['Incoming Stock'] > 0: return "Healthy (Incoming Planned)"
                    elif row['Avg. Weekly Demand'] == 0 and row['Total Expected Stock'] <= 0: return "Out of Stock & No Demand"
                    else: return "Unknown"
                    
                def get_trend_status(row):
                    if pd.isna(row['Change %']): return "Stable"
                    elif row['Change %'] > 0.3: return "Spike (>30%)"
                    elif row['Change %'] < -0.3: return "Drop (<-30%)"
                    else: return "Stable"

                sheet7 = df[df['In_Stock_Report'] == True].copy()
                noise_mask_7 = (sheet7['Current Stock'] < 1) & (sheet7['Avg. Weekly Demand'] == 0) & (sheet7['Incoming Stock'] == 0)
                sheet7 = sheet7[~noise_mask_7].copy()
                sheet7['Inventory Status'] = sheet7.apply(get_inv_status, axis=1)
                sheet7['Sales Trend'] = sheet7.apply(get_trend_status, axis=1)
                sheet7 = sheet7.sort_values(['Series', 'Brand', 'Code'])
                
                sheet8_raw = df[df['In_Stock_Report'] == True].groupby('Base SKU').agg({
                    'Series': 'first', 'Brand': 'first', 'Current Stock': 'sum', 'Incoming Stock': 'sum', 
                    'Overall Weekly Avg': 'sum', 'Weekly Prod Usage': 'sum', 'Current Week Sales': 'sum', 
                    'Prev Weekly Avg': 'sum', 'Quantity Change': 'sum'
                }).reset_index()
                
                sheet8_raw['Avg. Weekly Demand'] = sheet8_raw['Overall Weekly Avg'] + sheet8_raw['Weekly Prod Usage']
                sheet8_raw['Total Expected Stock'] = sheet8_raw['Current Stock'] + sheet8_raw['Incoming Stock']
                
                noise_mask_8 = (sheet8_raw['Current Stock'] < 1) & (sheet8_raw['Avg. Weekly Demand'] == 0) & (sheet8_raw['Incoming Stock'] == 0)
                sheet8_raw = sheet8_raw[~noise_mask_8].copy()
                
                sheet8_raw['Weeks of Supply'] = sheet8_raw.apply(calc_wos, axis=1)
                sheet8_raw['Change %'] = (sheet8_raw['Current Week Sales'] - sheet8_raw['Prev Weekly Avg']) / sheet8_raw['Prev Weekly Avg'].replace(0, pd.NA)
                sheet8_raw['Change % Display'] = (sheet8_raw['Change %'] * 100).fillna(0).round(1).astype(str) + "%"
                sheet8_raw['Inventory Status'] = sheet8_raw.apply(get_inv_status, axis=1)
                sheet8_raw['Sales Trend'] = sheet8_raw.apply(get_trend_status, axis=1)
                sheet8 = sheet8_raw.sort_values(['Series', 'Base SKU'])

                # --- 8. BULLETPROOF STYLING ENGINE ---
                def format_wos(val):
                    if pd.isna(val): return ""
                    if val >= 999: return "999+ weeks (No Demand)"
                    return f"{val:.2f} weeks"
                
                def format_demand(val):
                    if pd.isna(val): return ""
                    return f"{val:.2f}"

                def apply_styles(df_to_style, sheet_name):
                    # SAFETY: If the dataframe is empty, do not apply any styles to prevent math crashes.
                    if df_to_style.empty:
                        return df_to_style
                    
                    try:
                        styler = df_to_style.style
                        
                        # Number Formatters
                        format_dict = {}
                        if 'Weeks of Supply' in df_to_style.columns: format_dict['Weeks of Supply'] = format_wos
                        if 'Avg. Weekly Demand' in df_to_style.columns: format_dict['Avg. Weekly Demand'] = format_demand
                        if 'Quantity Change' in df_to_style.columns: format_dict['Quantity Change'] = '{:.2f}'
                        if 'Current Stock' in df_to_style.columns: format_dict['Current Stock'] = '{:.2f}'
                        styler = styler.format(format_dict)
                        
                        # Heatmaps (Gradients)
                        if sheet_name == "1. Dead Stock":
                            styler = styler.background_gradient(subset=['Current Stock'], cmap='Reds')
                        elif sheet_name == "2. Slow Movers":
                            styler = styler.background_gradient(subset=['Weeks of Supply'], cmap='Reds', vmin=10, vmax=52)
                        elif sheet_name == "3. Understock Risk":
                            styler = styler.background_gradient(subset=['Weeks of Supply'], cmap='Reds_r', vmin=0, vmax=4)
                            styler = styler.background_gradient(subset=['Avg. Weekly Demand'], cmap='Greens')
                        elif sheet_name == "4. Reorder Needed":
                            styler = styler.background_gradient(subset=['Weeks of Supply'], cmap='RdYlGn', vmin=4, vmax=10)
                        elif sheet_name == "5. Sales Spikes":
                            styler = styler.background_gradient(subset=['Quantity Change'], cmap='Greens')
                        elif sheet_name == "6. Sales Drops":
                            styler = styler.background_gradient(subset=['Quantity Change'], cmap='Reds_r')
                            
                        return styler
                    except Exception:
                        # Fallback: If any heatmap math fails, return raw table
                        return df_to_style
                
                # SAFETY: Function to safely write to Excel even if Python Styler formatting causes issues.
                def safe_write_excel(df, sheet_name, writer):
                    try:
                        styled = apply_styles(df, sheet_name)
                        styled.to_excel(writer, sheet_name=sheet_name, index=False)
                    except:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Prep the columns for display
                c1 = cat1[['Series', 'Code', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Action Recommended']]
                c2 = cat2[['Series', 'Code', 'Current Stock', 'Avg. Weekly Demand', 'Weeks of Supply', 'Action Recommended']]
                c3 = cat3[['Risk Level', 'Brand', 'Code', 'Weeks of Supply', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Avg. Weekly Demand']]
                c4 = cat4[['Brand', 'Code', 'Weeks of Supply', 'Current Stock', 'Avg. Weekly Demand']]
                c5 = cat5[['Series', 'Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']]
                c6 = cat6[['Series', 'Code', 'Quantity Change', 'Change % Display', 'Current Week Sales', 'Prev Weekly Avg', 'Current Stock']]
                s7_cols = ['Series', 'Brand', 'Base SKU', 'Code', 'Inventory Status', 'Sales Trend', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Avg. Weekly Demand', 'Weeks of Supply', 'Quantity Change', 'Change % Display']
                s8_cols = ['Series', 'Brand', 'Base SKU', 'Inventory Status', 'Sales Trend', 'Current Stock', 'Incoming Stock', 'Total Expected Stock', 'Avg. Weekly Demand', 'Weeks of Supply', 'Quantity Change', 'Change % Display']
                
                # --- 9. EXPORT TO EXCEL (SAFELY) ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    safe_write_excel(c1, "1. Dead Stock", writer)
                    safe_write_excel(c2, "2. Slow Movers", writer)
                    safe_write_excel(c3, "3. Understock Risk", writer)
                    safe_write_excel(c4, "4. Reorder Needed", writer)
                    safe_write_excel(c5, "5. Sales Spikes", writer)
                    safe_write_excel(c6, "6. Sales Drops", writer)
                    safe_write_excel(sheet7[s7_cols], "7. Master (Code)", writer)
                    safe_write_excel(sheet8[s8_cols], "8. Master (SKU)", writer)
                    
                    # Auto-width formatting
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
                st.success("Actionable items and Heatmaps generated successfully!")
                st.download_button(label="📥 Download Formatted Excel Report", data=buffer, file_name="Inventory_Action_Items.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                t1, t2, t3, t4, t5, t6, t7, t8 = st.tabs(["1. Dead", "2. Slow", "3. Understock", "4. Reorder", "5. Spikes", "6. Drops", "7. Master (Code)", "8. Master (SKU)"])
                
                with t1: st.dataframe(apply_styles(c1, "1. Dead Stock"), use_container_width=True)
                with t2: st.dataframe(apply_styles(c2, "2. Slow Movers"), use_container_width=True)
                with t3: st.dataframe(apply_styles(c3, "3. Understock Risk"), use_container_width=True)
                with t4: st.dataframe(apply_styles(c4, "4. Reorder Needed"), use_container_width=True)
                with t5: st.dataframe(apply_styles(c5, "5. Sales Spikes"), use_container_width=True)
                with t6: st.dataframe(apply_styles(c6, "6. Sales Drops"), use_container_width=True)
                with t7: st.dataframe(apply_styles(sheet7[s7_cols], "7. Master"), use_container_width=True)
                with t8: st.dataframe(apply_styles(sheet8[s8_cols], "8. Master"), use_container_width=True)

            except Exception as e:
                st.error(f"Error processing files: {e}")
