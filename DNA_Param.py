import streamlit as st
import pandas as pd
import re
import io
from io import StringIO, BytesIO

def process_multiple_datasets(file_content):
    datasets_raw = re.split(r'\n>|\n?>', file_content)
    if datasets_raw[0] == '' or not datasets_raw[0].strip():
        datasets_raw = datasets_raw[1:]
    elif datasets_raw[0].startswith('>'):
        datasets_raw[0] = datasets_raw[0][1:]

    results = []
    for dataset in datasets_raw:
        if not dataset.strip():
            continue
        sequence_id_match = re.match(r'(\d+\.\d+)', dataset)
        if not sequence_id_match:
            continue
        sequence_id = sequence_id_match.group(1)

        if "Intra-basepair parameters:" in dataset and "Inter-basepair parameters:" in dataset:
            parts = dataset.split("Inter-basepair parameters:")
            intra_part = parts[0]
            inter_part = "Inter-basepair parameters:" + parts[1]

            intra_lines = intra_part.split('\n')
            inter_lines = inter_part.split('\n')

            intra_header_line = next((line for line in intra_lines if "S.No" in line and "Basepair" in line), None)
            intra_data_lines = [line for line in intra_lines if re.match(r'^\s*\d+\s+[A-Z]{2}\s+', line)]

            inter_header_line = next((line for line in inter_lines if "S.No" in line and "BP step" in line), None)
            inter_data_lines = [line for line in inter_lines if re.match(r'^\s*\d+\s+[A-Z]{2}/[A-Z]{2}\s+', line)]

            if not intra_header_line or not intra_data_lines or not inter_header_line or not inter_data_lines:
                continue

            intra_headers = [h.strip() for h in intra_header_line.split('\t')]
            intra_data = [[val for val in re.split(r'\s+', line.strip())[:len(intra_headers)]] for line in intra_data_lines]

            inter_headers = [h.strip() for h in inter_header_line.split('\t')]
            inter_data = [[val for val in re.split(r'\s+', line.strip())[:len(inter_headers)]] for line in inter_data_lines]

            df_intra = pd.DataFrame(intra_data, columns=intra_headers)
            df_inter = pd.DataFrame(inter_data, columns=inter_headers)

            for df in [df_intra, df_inter]:
                for col in df.columns:
                    if col not in ["S.No", "Basepair", "BP step"]:
                        df[col] = pd.to_numeric(df[col], errors='coerce')

            sequence = "".join([bp[0] for bp in df_intra["Basepair"]])
            intra_avg = df_intra.drop(columns=["S.No", "Basepair"]).mean().round(5)
            inter_avg = df_inter.drop(columns=["S.No", "BP step"]).mean().round(5)

            result = pd.concat([
                pd.Series({"Sequence_ID": sequence_id, "Sequence": sequence}),
                inter_avg,
                intra_avg
            ]).to_frame().T

            results.append(result)

    return pd.concat(results, ignore_index=True) if results else None

st.title("🧬 cgDNA Excel Converter")

uploaded_file = st.file_uploader("Upload a DNA Parameter .txt file", type="txt")

if uploaded_file is not None:
    file_content = uploaded_file.read().decode("utf-8")
    results = process_multiple_datasets(file_content)
    
    if results is not None:
        # Extract base name and change extension to .xlsx
        base_filename = uploaded_file.name.rsplit(".", 1)[0]
        excel_filename = f"{base_filename}_processed.xlsx"
        
        # Convert to Excel in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            results.to_excel(writer, index=False)
        output.seek(0)

        st.success(f"Processed {len(results)} datasets.")
        st.download_button(
            label="Download Excel file",
            data=output,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

