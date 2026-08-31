import streamlit as st
import pandas as pd
import json
from io import BytesIO
from collections import OrderedDict

st.title("JSON to Excel Converter")

st.write(
    "Upload a JSON file (array of objects, or a single object). "
    "The Excel output mirrors the JSON exactly: same keys as columns, same "
    "values as cells, same order, same row count. Nothing is added, "
    "renamed, or recalculated."
)

uploaded_file = st.file_uploader("Upload JSON File", type=["json"])


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


if uploaded_file is not None:
    try:
        json_data = json.load(uploaded_file)
        records = to_records(json_data)

        # Preserve column order exactly as keys first appear in the JSON.
        columns = list(OrderedDict.fromkeys(key for record in records for key in record.keys()))

        rows = [
            {col: cell_value(record.get(col)) for col in columns}
            for record in records
        ]

        df = pd.DataFrame(rows, columns=columns)

        st.success(f"JSON Loaded Successfully! ({len(df)} rows, {len(df.columns)} columns)")
        st.dataframe(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except ValueError as e:
        st.error(str(e))
    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON file: {e}")
    except Exception as e:
        st.error(f"Error loading JSON: {e}")
