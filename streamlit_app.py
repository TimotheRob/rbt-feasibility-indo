import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="Formula Feasibility Checker", layout="wide")

st.title("🏭 Formula Production Feasibility Checker")

# --- SIDEBAR: FILE UPLOAD ---
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
    Recursively finds the BEST historical version of a product (Smart Explosion).
    Returns: (Set of Exploded Ingredients, BatchID Used, Availability Score)
    """
    if memo is None:
        memo = {}
    if path is None:
        path = set()

    # Check Cache
    if product_code in memo:
        return memo[product_code]

    # Check Circular Dependency
    if product_code in path:
        return (set(), "Circular Ref", 0.0)

    # If it's a raw material (no history of making it), return itself
    if product_code not in variants_map:
        is_in_stock = product_code in stock_set
        return ({product_code}, "Raw Material", 1.0 if is_in_stock else 0.0)

    # Retrieve all historical versions (batches) for this product
    # List of tuples: (BatchID, [IngredientCodes])
    possible_batches = variants_map[product_code]

    best_result = None
    best_score = -1.0

    path.add(product_code)

    # Evaluate every single historical batch to find the "Winner"
    for batch_id, ingredients in possible_batches:
        current_exploded_rms = set()

        # Explode this specific batch's ingredients
        valid_batch = True
        for ing in ingredients:
            exploded_ing_set, _, _ = get_best_recipe_path(
                ing, variants_map, stock_set, memo, path)
            # Optimization: If we hit a circular ref inside a batch, that batch is invalid, but we continue
            if exploded_ing_set == set():
                # Handle empty set logic if needed, but usually implies raw material or circular
                pass
            current_exploded_rms.update(exploded_ing_set)

        # Calculate Score for this Batch
        all_rms = list(current_exploded_rms)
        available_rms = [rm for rm in all_rms if rm in stock_set]

        if len(all_rms) > 0:
            ratio = len(available_rms) / len(all_rms)
        else:
            ratio = 0.0

        # Logic: Pick the one with higher availability.
        if ratio > best_score:
            best_score = ratio
            best_result = (current_exploded_rms, batch_id, ratio)
        elif ratio == best_score:
            # If tie, stick with current found or update if needed.
            # (First found is usually sufficient unless we have dates to break ties)
            if best_result is None:
                best_result = (current_exploded_rms, batch_id, ratio)

    path.remove(product_code)

    # If no valid batches found (rare), treat as missing item
    if best_result is None:
        best_result = ({product_code}, "No Valid Recipe", 0.0)

    memo[product_code] = best_result
    return best_result

# --- MAIN PROCESS ---


if target_file and stock_file and history_file:
    if st.button("Run Analysis"):
        with st.spinner("Processing..."):
            try:
                # 1. READ FILES

                # Target: A, B, C, D
                df_target = pd.read_excel(target_file, usecols="A:D", header=0)
                df_target.columns = ['Product Code',
                                     'Product Desc', 'Yearly Qty', '3M Qty']

                # Stock: D(3), I(8) -> (0-based)
                # But safer to use letters for clarity or verify
                # D=Raw Material Code, I=Raw Material Description
                df_stock = pd.read_excel(stock_file, usecols="D,I", header=0)
                df_stock.columns = ['RM Code', 'RM Desc']

                # History:
                # A(0) = RM Code
                # B(1) = RM Desc
                # D(3) = Manufacturing Order (Batch ID)  <-- FIXED as requested
                # J(9) = Product Code
                # K(10)= Product Desc
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

                # Create Description Dictionary
                desc_map = pd.Series(
                    df_history['RM Desc'].values, index=df_history['RM Code']).to_dict()
                desc_map.update(
                    pd.Series(df_stock['RM Desc'].values, index=df_stock['RM Code']).to_dict())

                # Available Stock Set
                stock_set = set(df_stock['RM Code'].unique())

                # 3. BUILD VARIANTS MAP (Grouping by Parent AND Batch)
                variants_map = {}
                grouped = df_history.groupby('Parent Code')

                for parent, group in grouped:
                    batch_groups = group.groupby('Batch ID')
                    variants_list = []
                    for batch_id, batch_rows in batch_groups:
                        ingredients = batch_rows['RM Code'].tolist()
                        variants_list.append((batch_id, ingredients))
                    variants_map[parent] = variants_list

                # 4. ANALYZE TARGETS
                results = []
                memo_cache = {}  # Cache for recursion

                prog_bar = st.progress(0)
                total_rows = len(df_target)

                for i, (index, row) in enumerate(df_target.iterrows()):
                    p_code = row['Product Code']

                    # Get Best Path
                    exploded_rms, best_batch_id, ratio = get_best_recipe_path(
                        p_code, variants_map, stock_set, memo_cache)

                    exploded_rms = list(exploded_rms)
                    num_exploded = len(exploded_rms)

                    available_rms = [
                        rm for rm in exploded_rms if rm in stock_set]
                    num_available = len(available_rms)

                    missing_rms = [
                        rm for rm in exploded_rms if rm not in stock_set]
                    num_missing = len(missing_rms)

                    # Missing String
                    missing_desc_list = []
                    for m_code in missing_rms:
                        desc = desc_map.get(m_code, "Unknown")
                        missing_desc_list.append(f"{m_code} ({desc})")
                    missing_str = "; ".join(missing_desc_list)

                    # Clean up Batch ID string for report
                    formula_ref = best_batch_id if best_batch_id != "Raw Material" else "N/A (Raw Material)"

                    results.append({
                        'Product Code': p_code,
                        'Product Description': row['Product Desc'],
                        'Yearly Qty': row['Yearly Qty'],
                        '3M Qty': row['3M Qty'],
                        'Best Formula Used (Batch)': formula_ref,
                        '# Ingredients (Exploded)': num_exploded,
                        '# Available': num_available,
                        'Availability Ratio': ratio,
                        '# Missing': num_missing,
                        'Missing List': missing_str
                    })

                    prog_bar.progress((i + 1) / total_rows)

                # 5. EXPORT
                df_result = pd.DataFrame(results)

                # Preview
                st.dataframe(df_result.head(10).style.format(
                    {'Availability Ratio': "{:.1%}"}))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_result.to_excel(writer, index=False,
                                       sheet_name='Analysis')

                    workbook = writer.book
                    worksheet = writer.sheets['Analysis']

                    # Formats
                    percent_fmt = workbook.add_format({'num_format': '0%'})
                    green_fmt = workbook.add_format(
                        {'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    red_fmt = workbook.add_format(
                        {'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    # Apply Percent Format to Column H (index 7)
                    worksheet.set_column('H:H', None, percent_fmt)

                    # Conditional Formatting
                    worksheet.conditional_format('H2:H1048576', {
                                                 'type': 'cell', 'criteria': '==', 'value': 1, 'format': green_fmt})
                    worksheet.conditional_format(
                        'H2:H1048576', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': red_fmt})

                buffer.seek(0)
                st.download_button(
                    label="📥 Download Excel Result",
                    data=buffer,
                    file_name="Formula_Availability_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.warning(
                    "Please check that the file columns match the required structure.")

else:
    st.info("Please upload all 3 files in the sidebar to begin.")
