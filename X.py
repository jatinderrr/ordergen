import pandas as pd

import numpy as np # Imported for rounding calculations

def style_and_write_sheet(df, writer, sheet_name):
    """
    Injects black separator rows between departments, styles them,
    and writes the result to a sheet in an ExcelWriter object.
    """
    # Ensure the dataframe is sorted by Department to group them correctly
    df_sorted = df.sort_values(by='Department')

    data_with_separators = []
    last_dept = None
    # Create a blank row that will be styled black
    separator_row = {col: '' for col in df_sorted.columns}

    for index, row in df_sorted.iterrows():
        current_dept = row['Department']
        # Add a separator if the department has changed (and it's not the first row)
        if last_dept is not None and current_dept != last_dept:
            data_with_separators.append(separator_row)
        data_with_separators.append(row.to_dict())
        last_dept = current_dept

    # If there's no data, don't create the sheet
    if not data_with_separators:
        return

    df_final = pd.DataFrame(data_with_separators)

    # Define the styling function
    def black_separator_styler(row):
        # If all cells in the row are empty, color the background black
        if all(cell == '' for cell in row):
            return ['background-color: black'] * len(row)
        else:
            return [''] * len(row)

    # Apply the style and write to the specified Excel sheet
    df_final.style.apply(black_separator_styler, axis=1).to_excel(
        writer, sheet_name=sheet_name, index=False
    )

#NEW FUNCTION 1

def write_plain_sheet(df, writer, sheet_name):
    if df.empty:
        return
    df.to_excel(writer, sheet_name=sheet_name, index=False)
#----------------------------------------------------------------------------------------------------------------------------

def calculate_reorder_quantities(
    sales_file='sales.xlsx',
    inventory_file='inventory.xlsx',
    ignore_file='ignore.xlsx',
    irc_file='IRC.xlsx',
    auto_export=False
):


    """
    Analyzes sales data and current inventory to calculate re-order quantities.
    """
    try:
        # --- Define Departments with Special 1-Week Supply Rule ---
        special_departments = [
            'BAKERY (WALL)- SHORT SHELF LIFE (BELOW 14 DAYS)',
            'BAKERY - COOLER',
            'BAKERY-SHORT SHELF LIFE (BELOW 14 DAYS)',
            'COOLER - CHEESE / BUTTER',
            'COOLER - DESSERT',
            'COOLER - DIP',
            'COOLER - MEAT',
            'COOLER - READY TO USE / EAT / DRINK',
            'YOGURT/YOGURT DRINK',
            'MILK'
        ]
        # Create a lowercase version for case-insensitive matching
        special_departments_lower = [d.lower() for d in special_departments]


        # --- Load Ignore List (Optional) ---
        try:
            df_ignore = pd.read_excel(ignore_file)
            ignore_codes = set(df_ignore['Stock Code'].astype(str))
            print(f"Info: Successfully loaded {len(ignore_codes)} stock codes from '{ignore_file}'.")
        except FileNotFoundError:
            print(f"Info: The ignore file '{ignore_file}' was not found.")
            ignore_codes = set()

        # --- Load and Process Inventory Data ---
        try:
            df_inventory = pd.read_excel(inventory_file)
            df_inventory['Stock Code'] = df_inventory['Stock Code'].astype(str)

            # Handle both "inventory.xlsx" and "stock analysis.xlsx"
            cols_lower = {c.lower(): c for c in df_inventory.columns}

            if 'quantity' in cols_lower:
                qty_col = cols_lower['quantity']          # inventory.xlsx
            elif 'qty. closing' in cols_lower:
                qty_col = cols_lower['qty. closing']      # stock analysis.xlsx
            else:
                raise KeyError("Could not find 'Quantity' or 'Qty. Closing' column in inventory file")

            df_inventory['Quantity'] = df_inventory[qty_col].clip(lower=0)

            # Extra info for IRC matching: description (col C) and department (col O)
            df_inventory['Inv Description'] = df_inventory.iloc[:, 2]
            df_inventory['Inv Department'] = df_inventory.iloc[:, 14]

            # Lookups
            inventory_lookup = df_inventory.set_index('Stock Code')['Quantity']
            inv_desc_lookup = df_inventory.set_index('Stock Code')['Inv Description']
            inv_dept_lookup = df_inventory.set_index('Stock Code')['Inv Department']

            inventory_loaded = True
        except FileNotFoundError:
            print(f"Warning: The inventory file '{inventory_file}' was not found. 'On Hand' quantities will be unknown.")
            inventory_loaded = False


        # --- Load and Process Sales Data ---
        df_sales = pd.read_excel(sales_file)

        # Normalize date column so we always have 'Stock Date'
        cols_lower = {c.lower(): c for c in df_sales.columns}

        if 'stock date' in cols_lower:
            date_col = cols_lower['stock date']          # sales.xlsx
        elif 'document date' in cols_lower:
            date_col = cols_lower['document date']       # sales detail.xlsx
        else:
            raise KeyError("Could not find 'Stock Date' or 'Document Date' in sales file")

        df_sales = df_sales.rename(columns={date_col: 'Stock Date'})

        # --- Filtering Rules ---
        df_sales['Stock Code'] = df_sales['Stock Code'].astype(str)


        if ignore_codes:
            df_sales = df_sales[~df_sales['Stock Code'].isin(ignore_codes)]
        
        departments_to_ignore = ['#OPENITEM', 'LOOSE ITEM', 'ECO FEE ADS', 'DEPOSIT', 'Dempsters Bread']
        departments_to_ignore_lower = [d.lower() for d in departments_to_ignore]
        df_sales['Description'] = df_sales['Description'].astype(str)
        df_filtered = df_sales[~df_sales['Description'].str.lower().isin(departments_to_ignore_lower)]

        desc_strings_to_ignore = ['MONDOUX', 'GREAT CANADIAN MEAT', 'LOOSE']
        df_filtered['Stock Description'] = df_filtered['Stock Description'].astype(str)
        df_filtered = df_filtered[~df_filtered['Stock Description'].str.contains('|'.join(desc_strings_to_ignore), case=False, na=False)]

        df_filtered = df_filtered[~df_filtered['Stock Code'].str.contains('LOOSE', case=False, na=False)]

        # --- Sales Calculations ---
        df_filtered['Stock Date'] = pd.to_datetime(df_filtered['Stock Date'])
        min_date = df_filtered['Stock Date'].min()
        max_date = df_filtered['Stock Date'].max()
        time_frame_days = (max_date - min_date).days if (max_date - min_date).days > 0 else 1

        product_sales = df_filtered.groupby(['Stock Code', 'Stock Description', 'Description'])['Quantity'].sum().reset_index()

        product_sales['Avg Daily Sales'] = product_sales['Quantity'] / time_frame_days
        product_sales['1 Week Sales'] = product_sales['Avg Daily Sales'] * 7
        product_sales['2 Week Sales'] = product_sales['Avg Daily Sales'] * 14
        product_sales['3 Week Sales'] = product_sales['Avg Daily Sales'] * 21
        product_sales['4 Week Sales'] = product_sales['Avg Daily Sales'] * 28

        # --- Ignore weekly sales figures that are less than 1 ---
        for i in range(1, 5):
            sales_col = f'{i} Week Sales'
            product_sales.loc[product_sales[sales_col] < 1, sales_col] = 0
        # --- Load IRC Data (Optional) ---
        try:
            df_irc_raw = pd.read_excel(irc_file)

            df_irc = pd.DataFrame({
                'Stock Code': df_irc_raw.iloc[:, 0].astype(str),  # Column A
                'DESCRIPTION': df_irc_raw.iloc[:, 1],             # Column B
                'IRC AMT': df_irc_raw.iloc[:, 3],                 # Column D
                'START DATE': df_irc_raw.iloc[:, 7],              # Column H
                'END DATE': df_irc_raw.iloc[:, 8]                 # Column I
            })


            df_irc = df_irc.dropna(subset=['Stock Code'])

            irc_lookup_amt = df_irc.set_index('Stock Code')['IRC AMT']
            irc_lookup_end = df_irc.set_index('Stock Code')['END DATE']
            irc_loaded = True

        except:
            irc_loaded = False


        # --- Merge Inventory Data ---
        if inventory_loaded:
            product_sales['On Hand'] = product_sales['Stock Code'].map(inventory_lookup)
            product_sales['On Hand'] = product_sales['On Hand'].fillna('INVENTORY UNKNOWN')
        else:
            product_sales['On Hand'] = 'INVENTORY UNKNOWN'
            
        # --- Merge IRC Data ---
        if irc_loaded:
            product_sales['IRC AMT'] = product_sales['Stock Code'].map(irc_lookup_amt)
            product_sales['END DATE'] = product_sales['Stock Code'].map(irc_lookup_end)
        else:
            product_sales['IRC AMT'] = ''
            product_sales['END DATE'] = ''


        # --- Print Final Output to Screen ---
        product_sales_sorted = product_sales.sort_values(by='Description')
        print(f"\nSales data analyzed from {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')} ({time_frame_days} days)\n")
        for index, row in product_sales_sorted.iterrows():
            print(f"Stock Code: {row['Stock Code']}", f"Product: {row['Stock Description']}", f"Department: {row['Description']}", sep='\n')
            on_hand_display = int(row['On Hand']) if isinstance(row['On Hand'], (int, float)) else row['On Hand']
            print(f"On Hand: {on_hand_display}")
            print(f"  - Sales per 1 week: {row['1 Week Sales']:.2f}", f"  - Sales per 2 weeks: {row['2 Week Sales']:.2f}", f"  - Sales per 3 weeks: {row['3 Week Sales']:.2f}", f"  - Sales per 4 weeks: {row['4 Week Sales']:.2f}", sep='\n')
            print("-" * 30)


        # --- Ask User to Generate Excel Report ---
                # --- Ask User to Generate Excel Report ---
        if auto_export:
            export_choice = 'yes'
        else:
            export_choice = input("\nWould you like to generate an Excel sheet with this data? (yes/no): ")

        if export_choice.lower().strip().startswith('y'):

            try:
                file_name = "reorder_report.xlsx"
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    # --- Sheet 1: FULL DATA ---
                    full_report_df = product_sales.rename(columns={'Stock Description': 'Product', 'Description': 'Department'})
                    full_report_df = full_report_df[['Department', 'Stock Code', 'Product', 'On Hand', 'IRC AMT', 'END DATE', '1 Week Sales', '2 Week Sales', '3 Week Sales', '4 Week Sales']]
                    style_and_write_sheet(full_report_df, writer, 'FULL DATA')

                    # --- Sheets 2-5: WEEKLY SUPPLY ---
                    numeric_on_hand = pd.to_numeric(product_sales['On Hand'], errors='coerce')
                    
                    is_special_dept_mask = product_sales['Description'].str.lower().isin(special_departments_lower)
                    special_items_to_order = product_sales[is_special_dept_mask]
                    ordered_codes = set()  # track items that are on any order sheet

                    for i in range(1, 5):
                        sales_col = f'{i} Week Sales'
                        
                        is_regular_dept_mask = ~product_sales['Description'].str.lower().isin(special_departments_lower)
                        needs_ordering_regular = (numeric_on_hand < product_sales[sales_col]) | (numeric_on_hand.isna())
                        sufficient_sales_regular = product_sales[sales_col] >= 1
                        regular_items_to_order = product_sales[is_regular_dept_mask & needs_ordering_regular & sufficient_sales_regular]
                        
                        supply_df = pd.concat([regular_items_to_order, special_items_to_order]).copy()
                        
                        if not supply_df.empty:
                            is_special_dept = supply_df['Description'].str.lower().isin(special_departments_lower)
                            on_hand_numeric_filtered = pd.to_numeric(supply_df['On Hand'], errors='coerce')
                            
                            effective_on_hand = np.where(is_special_dept, 0, on_hand_numeric_filtered)

                            base_sales = np.where(is_special_dept, supply_df['1 Week Sales'], supply_df[sales_col])
                            
                            qty_needed = base_sales - effective_on_hand
                            
                            qty_needed = np.where(on_hand_numeric_filtered.isna(), base_sales, qty_needed)
                            
                            supply_df['Quantity to Order'] = np.ceil(qty_needed)

                            supply_df = supply_df[supply_df['Quantity to Order'] > 0]

                            ordered_codes.update(supply_df['Stock Code'].astype(str).tolist())


                            if not supply_df.empty:
                                supply_df = supply_df.rename(columns={'Stock Description': 'Product', 'Description': 'Department'})
                                supply_df = supply_df[['Department', 'Stock Code', 'Product', 'On Hand', 'IRC AMT', 'END DATE', sales_col, 'Quantity to Order']]
                                style_and_write_sheet(supply_df, writer, f'{i} WEEKS SUPPLY')

                    # Lookups from sales for description + department
                    product_desc_lookup = product_sales.set_index('Stock Code')['Stock Description']
                    product_dept_lookup = product_sales.set_index('Stock Code')['Description']


                    # --- IRC SHEETS ---

                    if irc_loaded:
                        # Codes present in sales or inventory
                        sales_codes = set(product_sales['Stock Code'].astype(str))
                        if inventory_loaded:
                            inventory_codes = set(df_inventory['Stock Code'].astype(str))
                        else:
                            inventory_codes = set()
                        base_codes = sales_codes.union(inventory_codes)

                        irc_codes = set(df_irc['Stock Code'].astype(str))

                        # 1 IRC sheet: in IRC + (sales or inventory) BUT NOT on any order sheet
                        irc_existing_codes = irc_codes & base_codes
                        irc_not_ordered_codes = irc_existing_codes - ordered_codes

                        irc_existing_df = df_irc[df_irc['Stock Code'].isin(irc_not_ordered_codes)].copy()

                        if not irc_existing_df.empty:
                            # Bring in On Hand + week sales from product_sales
                            base_cols = product_sales[[
                                'Stock Code',
                                'On Hand',
                                '1 Week Sales',
                                '2 Week Sales',
                                '3 Week Sales',
                                '4 Week Sales'
                            ]]

                            irc_existing_merged = irc_existing_df.merge(
                                base_cols, on='Stock Code', how='left'
                            )

                            # Fill On Hand from inventory for items that never sold
                            if inventory_loaded:
                                irc_existing_merged['On Hand'] = irc_existing_merged['On Hand'].fillna(
                                    irc_existing_merged['Stock Code'].map(inventory_lookup)
                                )

                            # Week sales: if missing, treat as 0
                            for col in ['1 Week Sales', '2 Week Sales', '3 Week Sales', '4 Week Sales']:
                                irc_existing_merged[col] = irc_existing_merged[col].fillna(0)

                            # Build final description + department:
                            # 1st choice: sales file, 2nd: inventory file, 3rd: description from IRC list
                            irc_existing_merged['Desc_from_sales'] = irc_existing_merged['Stock Code'].map(product_desc_lookup)
                            irc_existing_merged['Dept_from_sales'] = irc_existing_merged['Stock Code'].map(product_dept_lookup)

                            if inventory_loaded:
                                irc_existing_merged['Desc_from_inv'] = irc_existing_merged['Stock Code'].map(inv_desc_lookup)
                                irc_existing_merged['Dept_from_inv'] = irc_existing_merged['Stock Code'].map(inv_dept_lookup)
                            else:
                                irc_existing_merged['Desc_from_inv'] = None
                                irc_existing_merged['Dept_from_inv'] = None

                            irc_existing_merged['Final_Description'] = (
                                irc_existing_merged['Desc_from_sales']
                                .combine_first(irc_existing_merged['Desc_from_inv'])
                                .combine_first(irc_existing_merged['DESCRIPTION'])
                            )

                            irc_existing_merged['Final_Department'] = (
                                irc_existing_merged['Dept_from_sales']
                                .combine_first(irc_existing_merged['Dept_from_inv'])
                            )

                            irc_sheet_df = irc_existing_merged[[
                                'Stock Code',
                                'Final_Description',
                                'On Hand',
                                '1 Week Sales',
                                '2 Week Sales',
                                '3 Week Sales',
                                '4 Week Sales',
                                'IRC AMT',
                                'END DATE',
                                'Final_Department'
                            ]].rename(columns={
                                'Final_Description': 'Description',
                                'Final_Department': 'Department'
                            })

                            #style_and_write_sheet(irc_sheet_df, writer, 'IRC')
                            write_plain_sheet(irc_sheet_df, writer, 'IRC')



                        # 2) IRC NEW ITEMS: only in IRC, not in sales or inventory
                        irc_new_codes = irc_codes - base_codes
                        irc_new_df = df_irc[df_irc['Stock Code'].isin(irc_new_codes)].copy()

                        if not irc_new_df.empty:
                            irc_new_sheet_df = irc_new_df[[
                                'Stock Code',
                                'DESCRIPTION',
                                'IRC AMT',
                                'START DATE',
                                'END DATE'
                            ]].rename(columns={'DESCRIPTION': 'Description'})

                            irc_new_sheet_df['Department'] = 'IRC NEW ITEMS'

                            style_and_write_sheet(irc_new_sheet_df, writer, 'IRC NEW ITEMS')





                    # --- NEW: Auto-fit column widths for all sheets ---
                    workbook = writer.book
                    for sheet in workbook.worksheets:
                        for column_cells in sheet.columns:
                            max_length = 0
                            column_letter = column_cells[0].column_letter
                            for cell in column_cells:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = (max_length + 2)
                            sheet.column_dimensions[column_letter].width = adjusted_width

                print(f"\n✅ Success! Formatted multi-sheet report saved as {file_name}")
            except Exception as e:
                print(f"\n❌ Error: Could not save the Excel file. Reason: {e}")
        else:
            print("\nNo file generated. Exiting script.")

    except FileNotFoundError as e:
        print(f"FATAL ERROR: A required file was not found: {e.filename}")
    except KeyError as e:
        print(f"FATAL ERROR: A required column is missing from a spreadsheet: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    calculate_reorder_quantities()