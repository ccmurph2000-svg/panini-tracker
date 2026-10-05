import os
import io
import pandas as pd
import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule

# Page configuration & Panini styling
st.set_page_config(page_title="Panini World Cup 2026 Tracker", page_icon="⚽", layout="centered")

st.title("⚽ Panini World Cup 2026 Album & Swap Tracker")
st.markdown("Upload photos of your World Cup 2026 album pages to track missing stickers, generate a styled Excel tracker with live charts, and share your swap list!")

# 1. API Configuration
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.text_input("Enter your Google Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    
    # 2. File Uploader for Album Pages
    uploaded_files = st.file_uploader(
        "Upload Panini World Cup 2026 Page Photos", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        st.subheader("📸 Uploaded Pages")
        cols = st.columns(len(uploaded_files))
        for idx, file in enumerate(uploaded_files):
            cols[idx].image(file, caption=f"Page {idx+1}", use_column_width=True)

        if st.button("🚀 Process Pages with AI", type="primary"):
            with st.spinner("Analyzing World Cup 2026 stickers, reading numbers, player names, and categories..."):
                try:
                    # Initialize updated Gemini Vision Model
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    # Prepare images for the API
                    image_parts = []
                    for file in uploaded_files:
                        bytes_data = file.getvalue()
                        image_parts.append({
                            'mime_type': file.type,
                            'data': bytes_data
                        })

                    prompt = """
                    Analyze these Panini World Cup 2026 sticker album pages. Extract all the visible slots or pasted stickers.
                    For each sticker, provide:
                    1. Section (e.g., 'Teams', 'Coca Cola / Special', 'Legends', 'Stadiums')
                    2. Team / Category (e.g., 'USA', 'France', 'Argentina', 'Coca Cola')
                    3. Sticker Number (e.g., 'USA 1', 'FWC 4', '15')
                    4. Player Name or Sticker Title (e.g., 'Kylian Mbappé', 'Official Mascot')
                    5. Status (Default to 'Need' for uncollected stickers, or 'Have' if it looks like a sticker is already pasted there).

                    Return ONLY a valid JSON list of objects with these keys: 
                    section, team, sticker_number, player_name, status
                    """

                    response = model.generate_content([prompt] + image_parts)
                    
                    import json
                    cleaned_text = response.text.strip()
                    if cleaned_text.startswith("```json"):
                        cleaned_text = cleaned_text[7:-3].strip()
                    
                    sticker_data = json.loads(cleaned_text)
                    df = pd.DataFrame(sticker_data)
                    
                    st.session_state['sticker_df'] = df
                    st.success("Successfully processed World Cup pages!")

                except Exception as e:
                    st.error(f"Error processing images: {e}")

    # 3. Managing Master List & Styled Excel Export
    if 'sticker_df' in st.session_state:
        st.divider()
        st.subheader("📋 Master Sticker Checklist")
        st.markdown("Review and update your collection status below (select **'Have'** or **'Need'**):")

        # Use selectbox options for the status column in Streamlit data editor
        edited_df = st.data_editor(
            st.session_state['sticker_df'], 
            num_rows="dynamic",
            column_config={
                "status": st.column_config.SelectboxColumn(
                    "Status",
                    help="Do you have this sticker or do you need it?",
                    options=["Have", "Need"],
                    required=True
                )
            }
        )

        def generate_panini_excel(df):
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='WC 2026 Collection')
            writer.close()
            
            output.seek(0)
            wb = load_workbook(output)
            ws = wb.active
            
            max_row = ws.max_row
            max_col = ws.max_column
            
            # --- Panini Branding & Formatting Styles ---
            header_fill = PatternFill(start_color="1A2B4C", end_color="1A2B4C", fill_type="solid") # Dark Panini Blue
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            
            collected_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid") # Soft mint green for "Have"
            collected_font = Font(name="Calibri", size=11, strike=True, color="595959")
            
            # Style Header
            for col in range(1, max_col + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Add Conditional Formatting: Strikethrough & Green Highlighting when cell equals "Have"
            have_rule = CellIsRule(operator='equal', formula=['"Have"'], stopIfTrue=True, font=collected_font, fill=collected_fill)
            ws.conditional_formatting.add(f"A2:E{max_row}", have_rule)

            # Auto-fit column widths
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            # Add Summary Table for Pie Chart counting "Have" and "Need"
            ws['G1'] = "Status"
            ws['H1'] = "Count"
            ws.cell(row=1, column=7).fill = header_fill
            ws.cell(row=1, column=7).font = header_font
            ws.cell(row=1, column=8).fill = header_fill
            ws.cell(row=1, column=8).font = header_font
            
            ws['G2'] = "Have"
            ws['H2'] = f'=COUNTIF(E2:E{max_row}, "Have")'
            ws['G3'] = "Need"
            ws['H3'] = f'=COUNTIF(E2:E{max_row}, "Need")'

            # Create Pie Chart
            pie = PieChart()
            pie.title = "Panini World Cup 2026 Completion"
            labels = Reference(ws, min_col=7, min_row=2, max_row=3)
            data = Reference(ws, min_col=8, min_row=1, max_row=3)
            pie.add_data(data, titles_from_data=True)
            pie.set_categories(labels)
            
            # Show percentages on chart labels
            pie.dataLabels = DataLabelList()
            pie.dataLabels.showPercent = True
            pie.dataLabels.showVal = False

            ws.add_chart(pie, "J2")
            
            final_output = io.BytesIO()
            wb.save(final_output)
            final_output.seek(0)
            return final_output

        excel_file = generate_panini_excel(edited_df)

        # Download Excel button
        st.download_button(
            label="📥 Download Panini World Cup 2026 Excel Tracker",
            data=excel_file,
            file_name="Panini_World_Cup_2026_Tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 4. WhatsApp & Email Sharing for Swaps (filtering by status == 'Need')
        st.divider()
        st.subheader("📤 Share Missing Stickers for Swaps")
        
        needed_df = edited_df[edited_df['status'] == 'Need']
        needed_text = "⚽ Hey! Here are the Panini World Cup 2026 stickers I still need to finish my album:\n\n"
        for index, row in needed_df.iterrows():
            needed_text += f"- [{row['team']}] Sticker #{row['sticker_number']}: {row['player_name']}\n"
        
        needed_text += "\nLet me know what you have for swaps!"

        col1, col2 = st.columns(2)
        
        with col1:
            import urllib.parse
            encoded_text = urllib.parse.quote(needed_text)
            whatsapp_url = f"[https://wa.me/?text=](https://wa.me/?text=){encoded_text}"
            st.link_button("💬 Share via WhatsApp", whatsapp_url)

        with col2:
            email_subject = urllib.parse.quote("Panini World Cup 2026 - Missing Stickers List")
            email_body = urllib.parse.quote(needed_text)
            email_url = f"mailto:?subject={email_subject}&body={email_body}"
            st.link_button("📧 Share via Email", email_url)

else:
    st.info("Please enter your Gemini API key above to start scanning your World Cup album pages.")
