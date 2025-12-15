import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="Formula Feasibility Checker", layout="wide")

st.title("🏭 Formula Production Feasibility Checker")

# --- SIDEBAR ---
st.sidebar.header("Upload Files")
target_file = st.sidebar.file_uploader(
    "1. Target Formulas (Excel)", type=['xlsx', 'xls'])
stock_file = st.sidebar.file_uploader(
    "2. Current Stock (Excel)", type=['xlsx', 'xls'])
history_file = st.sidebar.file_uploader(
    "3. Weighing History (Excel)", type=['xlsx', 'xls'])

# --- LOGIC FUNCTIONS ---


def normalize_code(series):
    """Ensures codes are treated as strings to avoid TypeErrors."""
    return series.astype(str).str.strip()


def get_best_recipe_path(product_code, variants_map, stock_set, memo=None, path=None):
    """
    Recursively finds the BEST historical version.
    Returns: (Set of Exploded Ingredients, BatchID Used, Availability Score, Missing Sources Dict)
    """
    if memo is None:
        memo = {}
    if path is None:
        path = set()

    # Check Cache
    if product_code in memo:
        return memo[product_code]

    # Check Circular
    if product_code in path:
        return (set(), "Circular Ref", 0.0, {})

    # Base Case: Raw Material (or Product with no history)
    if product_code not in variants_map:
        is_in_stock = product_code in stock_set
        # If missing, source is itself ("Direct")
        missing_src = {} if is_in_stock else {product_code: "Direct"}
        return ({product_code}, "Raw Material", 1.0 if is_in_stock else 0.0, missing_src)

    # Recursive Search
    possible_batches = variants_map[product_code]

    best_result = None
    best_score = -1.0
    best_len = -1

    path.add(product_code)

    for batch_id, ingredients in possible_batches:
        current_exploded_rms = set()
        current_missing_sources = {}

        # Explode this batch
        for ing in ingredients:
            exploded_set, _, _, child_missing_sources = get_best_recipe_path(
                ing, variants_map, stock_set, memo, path)

            current_exploded_rms.update(exploded_set)

            # Update Missing Sources tracking
            for missing_rm, source in child_missing_sources.items():
                if source == "Direct":
                    # The ingredient 'ing' is the missing item inside 'product_code'
                    current_missing_sources[missing_rm] = product_code
                else:
                    # Keep the deeper source (preserve the original parent)
                    current_missing_sources[missing_rm] = source

        # Calculate Score
        all_rms = list(current_exploded_rms)
        available_rms = [rm for rm in all_rms if rm in stock_set]
        total_count = len(all_rms)

        ratio = (len(available_rms) / total_count) if total_count > 0 else 0.0

        # Tie-Breaker Logic (Prefer higher ratio, then more complex formulas)
        is_better = False
        if ratio > best_score:
            is_better = True
        elif ratio == best_score:
            if total_count > best_len:
                is_better = True

        if is_better:
            best_score = ratio
            best_len = total_count
            best_result = (current_exploded_rms, batch_id,
                           ratio, current_missing_sources)

    path.remove(product_code)

    if best_result is None:
        missing_src = {product_code: "Direct"}
        best_result = ({product_code}, "No Valid Recipe", 0.0, missing_src)

    memo[product_code] = best_result
    return best_result

# --- MAIN PROCESS ---


if target_file and stock_file and history_file:
    if st.button("Run Analysis"):
        with st.spinner("Analyzing recursive structures..."):
            try:
                # 1. READ FILES
                df_target = pd.read_excel(target_file, usecols="A:D", header=0)
                df_target.columns = ['Product Code',
                                     'Product Desc', 'Yearly Qty', '3M Qty']

                df_stock = pd.read_excel(stock_file, usecols="D,I", header=0)
                df_stock.columns = ['RM Code', 'RM Desc']

                # History: A=RM, B=Desc, D=Batch, J=Parent, K=Desc
                df_history = pd.read_excel(history_file, usecols=[
                                           0, 1, 3, 9, 10], header=0)
                df_history.columns = ['RM Code', 'RM Desc',
                                      'Batch ID', 'Parent Code', 'Parent Desc']

                # 2. DATA CLEANING
                df_target['Product Code'] = normalize_code(
                    df_target['Product Code'])
                df_stock['RM Code'] = normalize_code(df_stock['RM Code'])
                df_history['RM Code'] = normalize_code(df_history['RM Code'])
                df_history['Parent Code'] = normalize_code(
                    df_history['Parent Code'])
                df_history['Batch ID'] = df_history['Batch ID'].astype(str)

                # Clean NaNs and Self-References
                df_history = df_history.dropna(
                    subset=['RM Code', 'Parent Code'])
                df_history = df_history[df_history['RM Code']
                                        != df_history['Parent Code']]

                # Build Dictionary for Descriptions
                # We need descriptions for Ingredients AND Intermediates (Parents)
                desc_map = pd.Series(
                    df_history['RM Desc'].values, index=df_history['RM Code']).to_dict()
                desc_map.update(
                    pd.Series(df_stock['RM Desc'].values, index=df_stock['RM Code']).to_dict())
                desc_map.update(pd.Series(
                    df_history['Parent Desc'].values, index=df_history['Parent Code']).to_dict())

                stock_set = set(df_stock['RM Code'].unique())

                # 3. BUILD VARIANTS MAP
                variants_map = {}
                grouped = df_history.groupby('Parent Code')
                for parent, group in grouped:
                    batch_groups = group.groupby('Batch ID')
                    variants_list = []
                    for batch_id, batch_rows in batch_groups:
                        ingredients = batch_rows['RM Code'].tolist()
                        if len(ingredients) > 0:
                            variants_list.append((batch_id, ingredients))
                    variants_map[parent] = variants_list

                # 4. ANALYZE TARGETS
                results = []
                memo_cache = {}
                prog_bar = st.progress(0)
                total_rows = len(df_target)

                for i, (index, row) in enumerate(df_target.iterrows()):
                    p_code = row['Product Code']

                    # RUN RECURSION
                    exploded_rms, best_batch_id, ratio, missing_sources = get_best_recipe_path(
                        p_code, variants_map, stock_set, memo_cache
                    )

                    exploded_rms = list(exploded_rms)
                    num_exploded = len(exploded_rms)
                    available_rms = [
                        rm for rm in exploded_rms if rm in stock_set]
                    num_available = len(available_rms)
                    num_missing = len(exploded_rms) - num_available

                    # GENERATE DETAILED MISSING LIST
                    missing_details = []
                    actual_missing_rms = [
                        rm for rm in exploded_rms if rm not in stock_set]

                    for m_code in actual_missing_rms:
                        m_desc = desc_map.get(m_code, "Unknown")

                        # Find where this came from
                        parent_code = missing_sources.get(m_code, "Direct")

                        if parent_code == p_code:
                            # Direct Ingredient
                            str_entry = f"{m_code} ({m_desc})"
                        else:
                            # Indirect Ingredient (Intermediate)
                            # Get description of the intermediate
                            p_desc = desc_map.get(
                                parent_code, "Unknown Intermediate")
                            # Format: Code (Desc) [via ParentCode - ParentDesc]
                            str_entry = f"{m_code} ({m_desc}) [via {parent_code} - {p_desc}]"

                        missing_details.append(str_entry)

                    # Changed Separator to just ";"
                    missing_str = ";".join(missing_details)
                    formula_ref = best_batch_id if best_batch_id != "Raw Material" else "N/A"

                    results.append({
                        'Product Code': p_code,
                        'Product Description': row['Product Desc'],
                        'Yearly Qty': row['Yearly Qty'],
                        '3M Qty': row['3M Qty'],
                        'Formula Used (Batch)': formula_ref,
                        '# Ingredients': num_exploded,
                        '# Available': num_available,
                        'Availability Ratio': ratio,
                        '# Missing': num_missing,
                        'Missing List': missing_str
                    })

                    prog_bar.progress((i + 1) / total_rows)

                # 5. EXPORT
                df_result = pd.DataFrame(results)

                st.dataframe(df_result.head(10).style.format(
                    {'Availability Ratio': "{:.1%}"}))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_result.to_excel(writer, index=False,
                                       sheet_name='Analysis')
                    workbook = writer.book
                    worksheet = writer.sheets['Analysis']

                    percent_fmt = workbook.add_format({'num_format': '0%'})
                    green_fmt = workbook.add_format(
                        {'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    red_fmt = workbook.add_format(
                        {'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    worksheet.set_column('H:H', None, percent_fmt)
                    worksheet.conditional_format(
                        'H2:H100000', {'type': 'cell', 'criteria': '==', 'value': 1, 'format': green_fmt})
                    worksheet.conditional_format(
                        'H2:H100000', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': red_fmt})

                buffer.seek(0)
                st.download_button(
                    label="📥 Download Detailed Analysis",
                    data=buffer,
                    file_name="Formula_Availability_Detailed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error: {e}")
                st.warning("Please ensure files match the column layout.")

else:
    st.info("Please upload files.")
