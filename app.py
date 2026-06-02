import streamlit as st
import pandas as pd

st.set_page_config(page_title="Inventory Predictor", page_icon="📦", layout="wide")
st.title("📦 Inventory Predictor & Analyzer")
st.write("Upload your Excel (.xlsx) files to instantly identify understock, overstock, and sales trends.")

# Create 4 columns for the 4 file uploads
col1, col2, col3, col4 = st.columns(4)
with col1:
    stock_file = st.file_uploader("1. Stock Level (.xlsx)", type=['xlsx'])
with col2:
    sales_file = st.file_uploader("2. Sales by SKU (.xlsx)", type=['xlsx'])
with col3:
    incoming_file = st.file_uploader("3. Incoming Shipments (.xlsx)", type=['xlsx'])
with col4:
    prod_file = st.file_uploader("4. Production Items (.xlsx)", type=['xlsx'])

# Run if the 3 core files are uploaded (Production file is optional but recommended)
if stock_file and sales_file and incoming_file:
    if st.button("Run Inventory Analysis", type="primary"):
        with st.spinner("Crunching weeks of supply and trends..."):
            
            try:
                # --- 1. PROCESS STOCK LEVEL ---
                stock_df = pd.read_excel(stock_file, sheet_name='Pivot Table', header=1)
                stock_grouped = stock_df.groupby('Product | Material Base SKU')['Total'].sum().reset_index()
                stock_grouped.rename(columns={'Total': 'Current Stock'}, inplace=True)

                # --- 2. PROCESS INCOMING SHIPMENTS ---
                incoming_df = pd.read_excel(incoming_file, sheet_name='Summary Data')
                incoming_grouped = incoming_df.groupby('Product | Material Base SKU')['Actual Outstanding Quantity'].sum().reset_index()
                incoming_grouped.rename(columns={'Actual Outstanding Quantity': 'Incoming Stock'}, inplace=True)

                # --- 3. PROCESS SALES & TRENDS ---
                sales_df = pd.read_excel(sales_file, sheet_name='Pivot Table', header=1)
                week_cols = [c for c in sales_df.columns if str(c).isdigit()]
                
                sales_grouped = sales_df.groupby('Product | Material Base SKU')[week_cols].sum().reset_index()
                sales_grouped['Weekly Avg Sales'] = sales_grouped[week_cols].mean(axis=1)
                
                if len(week_cols) >= 8:
                    sales_grouped['Recent 4W Sales'] = sales_grouped[week_cols[-4:]].sum(axis=1)
                    sales_grouped['Prev 4W Sales'] = sales_grouped[week_cols[-8:-4]].sum(axis=1)
                else:
                    half = len(week_cols) // 2
                    sales_grouped['Recent 4W Sales'] = sales_grouped[week_cols[-half:]].sum(axis=1)
                    sales_grouped['Prev 4W Sales'] = sales_grouped[week_cols[:half]].sum(axis=1)

                def get_trend(row):
                    if row['Prev 4W Sales'] == 0 and row['Recent 4W Sales'] > 0: return "📈 Sudden Spike"
                    if row['Prev 4W Sales'] == 0 and row['Recent 4W Sales'] == 0: return "➖ No Recent Sales"
                    change = (row['Recent 4W Sales'] - row['Prev 4W Sales']) / row['Prev 4W Sales']
                    if change >= 0.5: return f"📈 Spike (+{change:.0%})"
                    elif change <= -0.5: return f"📉 Drop ({change:.0%})"
                    else: return "➡️ Steady"

                sales_grouped['Sales Trend'] = sales_grouped.apply(get_trend, axis=1)

                # --- 4. PROCESS PRODUCTION ITEMS (IF UPLOADED) ---
                if prod_file:
                    prod_df = pd.read_excel(prod_file, sheet_name='Summary Data')
                    # Some items might be registered as 'In' or 'Out' on the factory floor, so we capture all movement
                    prod_df['Total Moved'] = prod_df.get('Quantity In', pd.Series(0)).fillna(0) + prod_df.get('Quantity Out', pd.Series(0)).fillna(0)
                    prod_grouped = prod_df.groupby('Product | Material Base SKU')['Total Moved'].sum().reset_index()
                    
                    # It's a 1-month report, so we divide by 4 to get the weekly production rate
                    prod_grouped['Weekly Prod Usage'] = prod_grouped['Total Moved'] / 4
                    prod_grouped.drop(columns=['Total Moved'], inplace=True)
                else:
                    prod_grouped = pd.DataFrame(columns=['Product | Material Base SKU', 'Weekly Prod Usage'])

                # --- 5. MERGE EVERYTHING TOGETHER ---
                all_skus = pd.DataFrame({'Product | Material Base SKU': pd.concat([
                    stock_grouped['Product | Material Base SKU'], 
                    incoming_grouped['Product | Material Base SKU'], 
                    sales_grouped['Product | Material Base SKU'],
                    prod_grouped['Product | Material Base SKU']
                ]).unique()})

                master_df = all_skus.merge(stock_grouped, on='Product | Material Base SKU', how='left')
                master_df = master_df.merge(incoming_grouped, on='Product | Material Base SKU', how='left')
                master_df = master_df.merge(sales_grouped[['Product | Material Base SKU', 'Weekly Avg Sales', 'Sales Trend']], on='Product | Material Base SKU', how='left')
                master_df = master_df.merge(prod_grouped, on='Product | Material Base SKU', how='left')

                # Fix blank data types cleanly
                master_df['Current Stock'] = master_df['Current Stock'].fillna(0)
                master_df['Incoming Stock'] = master_df['Incoming Stock'].fillna(0)
                master_df['Weekly Avg Sales'] = master_df['Weekly Avg Sales'].fillna(0)
                master_df['Weekly Prod Usage'] = master_df['Weekly Prod Usage'].fillna(0)
                master_df['Sales Trend'] = master_df['Sales Trend'].fillna("➖ No Recent Sales")

                # --- 6. CALCULATE WEEKS OF SUPPLY & CATEGORIZE ---
                master_df['Total Expected Stock'] = master_df['Current Stock'] + master_df['Incoming Stock']
                
                # Combine standard Sales and Internal Production for true demand
                master_df['Total Weekly Demand'] = master_df['Weekly Avg Sales'] + master_df['Weekly Prod Usage']
                
                def calc_wos(row):
                    if row['Total Weekly Demand'] <= 0:
                        return 999 if row['Total Expected Stock'] > 0 else 0
                    return row['Total Expected Stock'] / row['Total Weekly Demand']

                master_df['Weeks of Supply (WOS)'] = master_df.apply(calc_wos, axis=1)

                def categorize(row):
                    wos = row['Weeks of Supply (WOS)']
                    incoming = row['Incoming Stock']
                    
                    if wos < 4:
                        return "🚨 Understock (No Shipment)" if incoming == 0 else "⚠️ Understock (Shipment Arriving)"
                    elif wos > 10 and wos != 999:
                        return "📦 Overstock (No Shipment)" if incoming == 0 else "🛑 Overstock (Shipment Arriving!)"
                    elif wos == 999:
                        return "🧟 Dead Stock (No Movement)"
                    else:
                        return "✅ Healthy (4-10 Weeks)"

                master_df['Status'] = master_df.apply(categorize, axis=1)

                # Clean up formatting for the dashboard
                master_df['Weeks of Supply (WOS)'] = master_df['Weeks of Supply (WOS)'].apply(lambda x: "999+ (No Movement)" if x == 999 else round(x, 1))
                master_df['Total Weekly Demand'] = master_df['Total Weekly Demand'].round(2)
                master_df['Weekly Avg Sales'] = master_df['Weekly Avg Sales'].round(2)
                master_df['Weekly Prod Usage'] = master_df['Weekly Prod Usage'].round(2)
                
                # Reorder columns
                final_df = master_df[['Product | Material Base SKU', 'Status', 'Weeks of Supply (WOS)', 'Current Stock', 'Incoming Stock', 'Total Weekly Demand', 'Weekly Avg Sales', 'Weekly Prod Usage', 'Sales Trend']]

                # --- 7. DISPLAY DASHBOARD ---
                st.success("Analysis Complete!")
                
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "🚨 Understock (No Shipment)", 
                    "⚠️ Understock (Shipment Arriving)", 
                    "📦 Overstock (No Shipment)",
                    "🛑 Overstock (Shipment Arriving!)",
                    "📈 Significant Trends"
                ])
                
                with tab1: st.dataframe(final_df[final_df['Status'] == "🚨 Understock (No Shipment)"], use_container_width=True)
                with tab2: st.dataframe(final_df[final_df['Status'] == "⚠️ Understock (Shipment Arriving)"], use_container_width=True)
                with tab3: st.dataframe(final_df[final_df['Status'] == "📦 Overstock (No Shipment)"], use_container_width=True)
                with tab4: st.dataframe(final_df[final_df['Status'] == "🛑 Overstock (Shipment Arriving!)"], use_container_width=True)
                with tab5:
                    trend_df = final_df[final_df['Sales Trend'].str.contains("Spike|Drop")]
                    st.dataframe(trend_df, use_container_width=True)
                    
            except Exception as e:
                st.error(f"Error processing files: {e}")
                st.write("Please ensure you are uploading the correct .xlsx files and that they contain the standard sheets ('Pivot Table' and 'Summary Data').")
