import streamlit as st
import pandas as pd
import numpy as np
import time
import re
from io import BytesIO

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def highlight_total(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_1st_2nd_vendor(row, columns):
    styles = [""] * len(columns)
    first_vendor = row.get("1st Vendor")
    second_vendor = row.get("2nd Vendor")

    for i, col in enumerate(columns):
        if col == first_vendor:
            # styles[i] = "background-color: #f8c8dc; color: #7a1f47;"
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif col == second_vendor:
            # styles[i] = "background-color: #d7c6f3; color: #402e72;"
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"
    return styles

st.subheader("🧑‍🏫 User Guide: UPL Comparison")
st.markdown(
    ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
)
st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

# Divider custom
st.markdown(
    """
    <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="
        display: flex;
        align-items: center;
        height: 65px;
        margin-bottom: 10px;
    ">
        <div style="text-align: justify; font-size: 15px;">
            <span style="color: #C7FF00; font-weight: 800;">
            UPL Comparison</span>
            compares UPL values across multiple vendors in order to identify and 
            determine the most competitive item-level pricing.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The input file required for this menu should be a <span style="color: #FF69B4; font-weight: 500;">
            single file containing multiple sheets</span>, in eather <span style="background:#C6EFCE; 
            padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> 
            or <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">.xls</span> format. Each sheet represents a vendor name, with the 
            table structure in each sheet as follows:
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["Scope", "Desc", "Category", "UoM", "PRICE"]
df = pd.DataFrame([[""] * len(columns) for _ in range(3)], columns=columns)

st.dataframe(df, hide_index=True)

# Buat DataFrame 1 row
st.markdown("""
<table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
    <tr>
        <td style="border: 1px solid gray; width: 15%;">Vendor A</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor B</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor C</td>
        <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">multiple sheets</td>
    </tr>
</table>
""", unsafe_allow_html=True)

st.markdown("###### Description:")
st.markdown(
    """
    <div style="font-size:15px;">
        <ul>
            <li>
                <span style="display:inline-block; width:100px;">Scope - UoM</span>: non-numeric columns
            </li>
            <li>
                <span style="display:inline-block; width:100px;">PRICE</span>: numeric column
            </li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a <span style="font-weight: bold;">dynamic table</span>, 
            but it is <span style="color: #FF69B4; font-weight: 500;">ONLY APPLICABLE</span> 
            to <span style="color: #FF69B4; font-weight: 500;">non-numeric columns</span>.
            Unlike other menus, <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; 
            font-weight:600; font-size: 0.75rem; color: black">NUMERIC COLUMN</span> 
            is permitted <span style="color: #ED1C24; font-weight: bold;">ONLY ONCE</span> and <span 
            style="color: #ED1C24; font-weight: bold;">MUST</span> be placed in the last column.
            Also, users have the freedom to name the columns as they wish. The system logic relies on 
            <span style="font-weight: bold;">column indices</span>, not specific column names.
        </div>
    """,
    unsafe_allow_html=True 
)

st.markdown("**:violet-badge[Ensure that each sheet has the same table structure and column names!]**")

st.divider()
st.markdown("#### Constraint")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
            To ensure this menu works correctly, users need to follow certain rules regarding
            the dataset structure.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. COLUMN ORDER]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
            When creating tables, it is important to follow the specified column structure. Columns 
            <span style="font-weight: bold;">must</span> be arranged in the following order:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            Non-Numeric Columns → Numeric Column (only one)
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
            <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. NUMBER COLUMN]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Please refer the table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["No", "Scope", "Desc", "Category", "UoM", "PRICE"]
data = [
    [1] + [""] * (len(columns) - 1),
    [2] + [""] * (len(columns) - 1),
    [3] + [""] * (len(columns) - 1)
]
df = pd.DataFrame(data, columns=columns)

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
            The table above is an <span style="color: #FF69B4; font-weight: 700;">incorrect example</span>
            and is <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains 
            a <span style="font-weight: bold;">"No"</span> column. The "No" column is prohibited in this
            menu, as it will be treated as a numeric column by the system, which violates the constraint
            described in point 1.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables <span style="color: #FF69B4; font-weight: 700;">
            do not need to start from cell A1</span>. However, ensure
            that the cells above and to the left of the table are empty, as shown in the example below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["", "A", "B", "C", "D", "E", "F"]

# Buat 5 baris kosong
df = pd.DataFrame([[""] * len(columns) for _ in range(6)], columns=columns)

# Isi kolom pertama dengan 1–6
df.iloc[:, 0] = [1, 2, 3, 4, 5, 6]

# Header bagian kedua
df.loc[1, ["B", "C", "D", "E"]] = ["Desc", "Category", "UoM", "PRICE"]

# Data Software & Hardware
df.loc[2, ["B", "C", "D", "E"]] = ["Optical Cable", "Non-Services Area & Material", "M", "3.600"]
df.loc[3, ["B", "C", "D", "E"]] = ["Cross Connect", "Non-Services Area & Material", "Link", "29.800"]
df.loc[4, ["B", "C", "D", "E"]] = ["Dismantle RAU", "Non-Services Area & Material", "Pcs", "274.450"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top:-10px;">
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

st.divider()

st.markdown("#### What is Displayed?")

# Path file Excel yang sudah ada
file_path = "dummy dataset.xlsx"

# Buka file sebagai binary
with open(file_path, "rb") as f:
    file_data = f.read()

# Markdown teks
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
        You can try this menu by downloading the dummy dataset using the button below: 
    </div>
    """,
    unsafe_allow_html=True
)

@st.fragment
def release_the_balloons():
    st.balloons()

# Download button untuk file Excel
st.download_button(
    label="Dummy Dataset",
    data=file_data,
    file_name="dummy dataset.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    on_click=release_the_balloons,
    type="primary",
    use_container_width=True,
)

st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        Based on this dummy dataset, the menu will produce the following results.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MERGE DATA]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will merge the tables from each sheet into a single table and add
            a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">TOTAL ROW</span> for each vendor, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "Desc", "Category", "UoM", "Price (IDR)"]
data = [
    ["Vendor A", "Instalasi Optical Cable", "Non-Services Area & Material", "M", 3500],
    ["Vendor A", "Cross connect", "Non-Services Area & Material", "Link", 30000],
    ["Vendor A", "Dismantle RAU", "Non-Services Area & Material", "Pcs", 274500],
    ["Vendor A", "TOTAL", "", "", 308000],

    ["Vendor B", "Instalasi Optical Cable", "Non-Services Area & Material", "M", 3350],
    ["Vendor B", "Cross connect", "Non-Services Area & Material", "Link", 30800],
    ["Vendor B", "Dismantle RAU", "Non-Services Area & Material", "Pcs", 274200],
    ["Vendor B", "TOTAL", "", "", 308350],

    ["Vendor C", "Instalasi Optical Cable", "Non-Services Area & Material", "M", 3600],
    ["Vendor C", "Cross connect", "Non-Services Area & Material", "Link", 29800],
    ["Vendor C", "Dismantle RAU", "Non-Services Area & Material", "Pcs", 274450],
    ["Vendor C", "TOTAL", "", "", 307850],
]
df_merge = pd.DataFrame(data, columns=columns)

num_cols = ["Price (IDR)"]
df_merge_styled = (
    df_merge.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_merge_styled, hide_index=True)

st.write("")
st.markdown("**:orange-badge[2. TRANSPOSE DATA]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            After merging the data, the system will transpose the <span style="color: #FF69B4; 
            font-weight: 700;">VENDOR</span> column and add a TOTAL row at the bottom, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Desc", "Category", "UoM", "Vendor A", "Vendor B", "Vendor C"]
data = [
    ["Corss connect", "Non-Services Area & Material", "Link", 30000, 30800, 29800],
    ["Dismantle RAU", "Non-Services Area & Material", "Pcs", 274500, 274200, 274450],
    ["Instalasi Optical Cable", "Non-Services Area & Material", "M", 3500, 3350, 3600],
    ["TOTAL", "", "", 308000, 308350, 307850]
]
df_transpose = pd.DataFrame(data, columns=columns)

num_cols = ["Vendor A", "Vendor B", "Vendor C"]
df_transpose_styled = (
    df_transpose.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)
st.dataframe(df_transpose_styled, hide_index=True)

st.write("")
st.markdown("**:yellow-badge[3. BID & PRICE ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu also displays an analysis table that provides a comprehensive overview of the pricing structure 
            submitted by each vendor, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="text-align:left; margin-bottom: 8px">
        <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
        &nbsp;
        <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
    </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Desc", "Category", "UoM", "Vendor A", "Vendor B", "Vendor C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "Vendor A to Median (%)", "Vendor B to Median (%)", "Vendor C to Median (%)"]
data = [
    ["Cross connect", "Non-Services Area & Material", "Link", 30000, 30800, 29800, 29800, "Vendor C", 30000, "Vendor A", "0.7%", 30000, "+0.0%", "+2.7%", "-0.7%"],
    ["Dismantle RAU", "Non-Services Area & Material", "Pcs", 247500, 247200, 274450, 247200, "Vendor B", 274450, "Vendor C", "0.1%", 274450, "+0.0%", "-0.1%", "+0.0%"],
    ["Instalasi Optical Cable", "Non-Services Area & Material", "M", 3500, 3350, 3600, 3350, "Vendor B", 3500, "Vendor A", "4.5%", 3500, "+0.0%", "-4.3%", "+2.9%"],
]
df_analysis = pd.DataFrame(data, columns=columns)

num_cols = ["Vendor A", "Vendor B", "Vendor C", "1st Lowest", "2nd Lowest", "Median Price"]
df_analysis_styled = (
    df_analysis.style
    .format({col: format_rupiah for col in num_cols})
    .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis.columns), axis=1)
)

st.dataframe(df_analysis_styled, hide_index=True)

st.write("")
st.markdown("**:green-badge[4. VISUALIZATION]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu displays visualizations focusing on two key aspects: <span style="background: #FF5E5E; 
            padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Win Rate Trend</span> 
            and <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 13px; color: black">Average Gap Trend</span>, each presented in its own tab.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

with tab1:
    st.image("assets/1.png")
    with st.expander("See explanation"):
        st.caption('''
            The visualization above compares the win rate of each vendor
            based on how often they achieved 1st or 2nd place in all
            tender evaluations.  
                    
            **💡 How to interpret the chart**  
                    
            - High 1st Win Rate (%)  
                Vendor is highly competitive and often offers the best commercial terms.  
            - High 2nd Win Rate (%)  
                Vendor consistently performs well, often just slightly less competitive than the winner.  
            - Large Gap Between 1st & 2nd Win Rate  
                Shows clear market dominance by certain vendors.
        ''')

with tab2:
    st.image("assets/2.png")
    with st.expander("See explanation"):
        st.caption('''
            The chart above shows the average price difference between 
            the lowest and second-lowest bids for each vendor when they 
            rank 1st, indicating their pricing dominance or competitiveness.
                    
            **💡 How to interpret the chart**  
                    
            - High Gap  
                High gap indicates strong vendor dominance (much lower prices).  
            - Low Gap  
                Low gap indicates intense competition with similar pricing among vendors.  
            
            The dashed line represents the average gap across all vendors, serving as a benchmark (1.5%).
        ''')
    
st.write("")
st.markdown("**:blue-badge[5. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">Super Button</span> feature where all dataframes generated by the system 
            can be downloaded as a single file with multiple sheets. You can also customize the order of the sheets.
            The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

dataframes = {
    "Merge Data": df_merge,
    "Transpose Data": df_transpose,
    "Bid & Price Analysis": df_analysis,
}

# Tampilkan multiselect
selected_sheets = st.multiselect(
    "Select sheets to download in a single Excel file:",
    options=list(dataframes.keys()),
    default=list(dataframes.keys())  # default semua dipilih
)

# Fungsi "Super Button" & Formatting
def generate_multi_sheet_excel(selected_sheets, df_dict):
    """
    Buat Excel multi-sheet dengan highlight:
    - Sheet 'Bid & Price Analysis' -> highlight 1st & 2nd vendor
    - Sheet lainnya -> highlight row TOTAL
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet in selected_sheets:
            df = df_dict[sheet].copy()
            df.to_excel(writer, index=False, sheet_name=sheet)
            workbook  = writer.book
            worksheet = writer.sheets[sheet]

            # --- Format umum ---
            fmt_rupiah = workbook.add_format({'num_format': '#,##0'})
            fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})
            fmt_total  = workbook.add_format({
                "bold": True, "bg_color": "#D9EAD3", "font_color": "#1A5E20", "num_format": "#,##0"
            })
            fmt_first  = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})
            fmt_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"})

            # Identifikasi numeric columns
            numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
            vendor_cols = [c for c in numeric_cols] if sheet == "Bid & Price Analysis" else []

            # Apply format kolom numeric / persen
            for col_idx, col_name in enumerate(df.columns):
                if col_name in numeric_cols:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                if "%" in col_name:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

            # --- Highlight baris ---
            for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                # Cek apakah TOTAL
                is_total_row = any(str(x).strip().upper() == "TOTAL" for x in row if pd.notna(x))

                # Ambil nama 1st & 2nd vendor untuk sheet Bid & Price Analysis
                if sheet == "Bid & Price Analysis":
                    first_vendor_name = row[df.columns.get_loc("1st Vendor")]
                    second_vendor_name = row[df.columns.get_loc("2nd Vendor")]

                    # Cari index kolom vendor di vendor_cols
                    first_idx = df.columns.get_loc(first_vendor_name) if first_vendor_name in vendor_cols else None
                    second_idx = df.columns.get_loc(second_vendor_name) if second_vendor_name in vendor_cols else None

                # Loop tiap kolom
                for col_idx, col_name in enumerate(df.columns):
                    value = row[col_idx]
                    fmt = None

                    # Highlight TOTAL untuk sheet selain Bid & Price Analysis
                    if is_total_row and sheet in ["Merge Data", "Transpose Data"]:
                        fmt = fmt_total

                    # Highlight 1st/2nd vendor
                    elif sheet == "Bid & Price Analysis":
                        if first_idx is not None and col_idx == first_idx:
                            fmt = fmt_first
                        elif second_idx is not None and col_idx == second_idx:
                            fmt = fmt_second

                    # Tangani NaN / None / inf
                    if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                        value = ""

                    worksheet.write(row_idx, col_idx, value, fmt)

    output.seek(0)
    return output

# ---- DOWNLOAD BUTTON ----
if selected_sheets:
    excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

    st.download_button(
        label="Download",
        data=excel_bytes,
        file_name="super botton.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
        
st.write("")
st.divider()

st.markdown("#### Video Tutorial")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">YouTube</span> link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/gcgWa76-qpY?si=A7RDLRCSJd0fZst2")