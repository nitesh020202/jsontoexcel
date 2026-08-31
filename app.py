import streamlit as st
import pandas as pd
import json
import zipfile
from io import BytesIO
from collections import OrderedDict

st.title("JSON to Excel Converter")

st.write(
    "Upload one or more JSON files (array of objects, or a single object). "
    "Each file is converted independently and the Excel output mirrors its "
    "JSON exactly: same keys as columns, same values as cells, same order, "
    "same row count. Nothing is added, renamed, or recalculated."
)

uploaded_files = st.file_uploader(
    "Upload JSON File(s)", type=["json"], accept_multiple_files=True
)


def to_records(json_data):
    """Normalize input JSON into a list of dict records without altering values."""
    if isinstance(json_data, list):
        if not all(isinstance(item, dict) for item in json_data):
            raise ValueError(
                "Every element of the JSON array must be an object ({...}). "
                "Example: [{...}, {...}]"
            )
        return json_data
    if isinstance(json_data, dict):
        return [json_data]
    raise ValueError(
        "JSON root must be an object ({...}) or an array of objects ([{...}, ...])."
    )


def cell_value(value):
    """Keep scalars as-is; serialize nested structures verbatim as JSON text
    so no new columns are introduced and no data is lost or changed."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False)


def convert_to_excel_bytes(json_data):
    """Convert parsed JSON data into an in-memory Excel file. Returns
    (excel_bytes, row_count, column_count)."""
    records = to_records(json_data)

    # Preserve column order exactly as keys first appear in the JSON.
    columns = list(OrderedDict.fromkeys(key for record in records for key in record.keys()))

    rows = [
        {col: cell_value(record.get(col)) for col in columns}
        for record in records
    ]

    df = pd.DataFrame(rows, columns=columns)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

    return output.getvalue(), df


if uploaded_files:
    results = []  # (input_name, output_filename, excel_bytes) for successful conversions

    for uploaded_file in uploaded_files:
        st.subheader(uploaded_file.name)
        try:
            json_data = json.load(uploaded_file)
            excel_bytes, df = convert_to_excel_bytes(json_data)
            output_filename = f"{uploaded_file.name.rsplit('.', 1)[0]}.xlsx"

            st.success(f"Loaded successfully! ({len(df)} rows, {len(df.columns)} columns)")
            st.dataframe(df)

            st.download_button(
                label=f"📥 Download {output_filename}",
                data=excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{uploaded_file.name}_{uploaded_file.file_id}",
            )

            results.append((uploaded_file.name, output_filename, excel_bytes))

        except ValueError as e:
            st.error(str(e))
        except json.JSONDecodeError as e:
            st.error(f"Invalid JSON file: {e}")
        except Exception as e:
            st.error(f"Error loading JSON: {e}")

    if len(results) > 1:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, output_filename, excel_bytes in results:
                zf.writestr(output_filename, excel_bytes)

        st.divider()
        st.download_button(
            label="📦 Download All as ZIP",
            data=zip_buffer.getvalue(),
            file_name="converted_files.zip",
            mime="application/zip",
        )
