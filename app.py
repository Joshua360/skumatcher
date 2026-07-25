import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRODUCT_FILE = os.path.join(BASE_DIR, "AllProductsList.xlsx")


st.set_page_config(
    page_title="Invoice SKU Matcher",
    page_icon="📦",
    layout="wide"
)


st.title("📦 Invoice SKU Matcher")
st.write("Upload your invoice records file. The app will match each invoice line to the correct FullSKU.")


@st.cache_data
def load_product_data():
    product_data = pd.read_excel(PRODUCT_FILE)

    product_data.columns = product_data.columns.str.strip()

    required_columns = ["BaseDescription", "FullSKU"]
    missing_columns = [col for col in required_columns if col not in product_data.columns]

    if missing_columns:
        raise ValueError(f"Product file is missing required columns: {missing_columns}")

    product_data["BaseDescription"] = product_data["BaseDescription"].astype(str).str.strip()
    product_data["FullSKU"] = product_data["FullSKU"].astype(str).str.strip().str.upper()

    return product_data


def has_vintage_year(description):
    return re.search(r"\b(19|20)\d{2}\b", str(description)) is not None


def get_vintage_year(description):
    match = re.search(r"\b(19|20)\d{2}\b", str(description))
    return match.group(0) if match else None


def remove_vintage_year(description):
    return re.sub(r"\b(19|20)\d{2}\b", "", str(description)).strip()


def remove_last_word(description):
    parts = str(description).split()

    if len(parts) <= 1:
        return str(description), ""

    removed_value = parts[-1]
    new_description = " ".join(parts[:-1])

    return new_description, removed_value


def sku_prefix(fullsku):
    return str(fullsku).split("-")[0].strip().upper()


def search_products(search_text, product_data):
    search_text = str(search_text).strip()

    if not search_text:
        return product_data.iloc[0:0]

    return product_data[
        product_data["BaseDescription"].str.contains(
            re.escape(search_text),
            case=False,
            na=False
        )
    ]


def choose_by_invoice_code(matches, invoice_code):
    invoice_code = str(invoice_code).strip().upper()

    filtered = matches[
        matches["FullSKU"].apply(lambda x: sku_prefix(x) == invoice_code)
    ]

    if len(filtered) >= 1:
        return filtered.iloc[0]["FullSKU"]

    return None


def choose_by_invoice_code_and_value(matches, invoice_code, required_value):
    invoice_code = str(invoice_code).strip().upper()
    required_value = str(required_value).strip().upper()

    if not required_value:
        return None

    filtered = matches[
        (matches["FullSKU"].apply(lambda x: sku_prefix(x) == invoice_code)) &
        (matches["FullSKU"].str.upper().str.contains(re.escape(required_value), na=False))
    ]

    if len(filtered) >= 1:
        return filtered.iloc[0]["FullSKU"]

    return None


def match_invoice_line(row, product_data):
    invoice_desc = str(row["InvoiceDescription"]).strip()
    invoice_code = str(row["InvoiceCode"]).strip().upper()

    # 1. Search description exactly as it is
    matches = search_products(invoice_desc, product_data)

    if len(matches) == 1:
        return matches.iloc[0]["FullSKU"]

    if len(matches) > 1:
        selected = choose_by_invoice_code(matches, invoice_code)
        if selected:
            return selected

    # 2. If no result and description has vintage year
    if len(matches) == 0 and has_vintage_year(invoice_desc):
        vintage = get_vintage_year(invoice_desc)
        desc_without_vintage = remove_vintage_year(invoice_desc)

        matches = search_products(desc_without_vintage, product_data)

        if len(matches) == 1:
            return matches.iloc[0]["FullSKU"]

        if len(matches) > 1:
            selected = choose_by_invoice_code_and_value(matches, invoice_code, vintage)
            if selected:
                return selected

            selected = choose_by_invoice_code(matches, invoice_code)
            if selected:
                return selected

    # 3. If no result and description does not have vintage year
    if len(matches) == 0 and not has_vintage_year(invoice_desc):
        desc_without_last_word, removed_value = remove_last_word(invoice_desc)

        matches = search_products(desc_without_last_word, product_data)

        if len(matches) == 1:
            return matches.iloc[0]["FullSKU"]

        if len(matches) > 1:
            selected = choose_by_invoice_code_and_value(matches, invoice_code, removed_value)
            if selected:
                return selected

            selected = choose_by_invoice_code(matches, invoice_code)
            if selected:
                return selected

    return None


def extract_year_from_text(text):
    text = str(text)
    match = re.search(r"\b(19|20)\d{2}\b", text)
    return match.group(0) if match else ""


def extract_year_from_sku(fullsku):
    fullsku = str(fullsku)
    match = re.search(r"\b(19|20)\d{2}\b", fullsku)
    return match.group(0) if match else ""


def audit_match(row):
    invoice_desc = str(row.get("InvoiceDescription", "")).strip()
    matched_sku = str(row.get("MatchedFullSKU", "")).strip().upper()

    invoice_vintage = extract_year_from_text(invoice_desc)
    sku_vintage = extract_year_from_sku(matched_sku)

    # Only warn when invoice description has a vintage year,
    # matched SKU also has a vintage year,
    # and the two years are different.
    if invoice_vintage and sku_vintage and invoice_vintage != sku_vintage:
        return f"Vintage mismatch: invoice has {invoice_vintage} but matched SKU has {sku_vintage}"

    return ""


def convert_df_to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Matched Invoice Records")

    return output.getvalue()


try:
    product_data = load_product_data()
    st.success(f"Product list loaded successfully: {PRODUCT_FILE}")
except FileNotFoundError:
    st.error(f"Could not find {PRODUCT_FILE}. Make sure it is in the same folder as app.py.")
    st.stop()
except Exception as e:
    st.error(f"Error loading product file: {e}")
    st.stop()


uploaded_file = st.file_uploader(
    "Upload invoice records Excel file",
    type=["xlsx"]
)


if uploaded_file is not None:
    try:
        invoice_data = pd.read_excel(uploaded_file)

        invoice_data.columns = invoice_data.columns.str.strip()

        required_columns = ["InvoiceDescription", "InvoiceCode"]
        missing_columns = [col for col in required_columns if col not in invoice_data.columns]

        if missing_columns:
            st.error(f"Missing required columns in uploaded invoice file: {missing_columns}")
            st.stop()

        invoice_data["InvoiceDescription"] = invoice_data["InvoiceDescription"].astype(str).str.strip()
        invoice_data["InvoiceCode"] = invoice_data["InvoiceCode"].astype(str).str.strip().str.upper()

        st.subheader("Uploaded Invoice Preview")
        st.dataframe(invoice_data.head(20), use_container_width=True)

        if st.button("Match FullSKUs"):
            with st.spinner("Matching invoice records..."):
                invoice_data["MatchedFullSKU"] = invoice_data.apply(
                    match_invoice_line,
                    axis=1,
                    product_data=product_data
                )

                # Audit columns for identifying only true vintage mismatches
                invoice_data["InvoiceVintage"] = invoice_data["InvoiceDescription"].apply(extract_year_from_text)
                invoice_data["MatchedSKUVintage"] = invoice_data["MatchedFullSKU"].apply(extract_year_from_sku)
                invoice_data["MatchWarning"] = invoice_data.apply(audit_match, axis=1)

                total_records = len(invoice_data)
                matched_records = invoice_data["MatchedFullSKU"].notna().sum()
                unmatched_records = invoice_data["MatchedFullSKU"].isna().sum()

                suspicious_matches = invoice_data[
                    invoice_data["MatchWarning"].astype(str).str.strip() != ""
                ]

                suspicious_count = len(suspicious_matches)

                st.success("Matching complete.")

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total Records", total_records)
                col2.metric("Matched Records", matched_records)
                col3.metric("Unmatched Records", unmatched_records)
                col4.metric("Vintage Warnings", suspicious_count)

                if suspicious_count > 0:
                    st.warning(
                        f"{suspicious_count} vintage mismatch(es) found. "
                        "Review these before using the final file."
                    )

                    st.subheader("Vintage Mismatch Warnings")
                    st.dataframe(suspicious_matches, use_container_width=True)
                else:
                    st.success("No vintage mismatches found.")

                st.subheader("Matched Results Preview")
                st.dataframe(invoice_data.head(50), use_container_width=True)

                unmatched_data = invoice_data[invoice_data["MatchedFullSKU"].isna()]

                if unmatched_records > 0:
                    st.subheader("Unmatched Records")
                    st.dataframe(unmatched_data, use_container_width=True)

                matched_excel = convert_df_to_excel(invoice_data)

                st.download_button(
                    label="Download Matched Excel File",
                    data=matched_excel,
                    file_name="matched_invoice_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
        