import pandas as pd
import re
from pathlib import Path
import streamlit as st

st.title("Manifest Master Dashboard")

# === CONFIG ===
#input_path = Path("/content/drive/MyDrive/Manifest Lock/OD_Hub_Mapping.xlsx")
#output_path = Path("/content/drive/MyDrive/Manifest Lock/Final_Manifest_Rules_2_with mode.csv")

# === READ OD HUB MAPPING ===
hub_mapping_flag = st.radio(
        "Has Hub Mapping Changed?",
        ["No", "Yes", "Maybe, I donno"],

    )


if hub_mapping_flag == "Yes": 
        od_df = pd.read_excel("OD_Hub_Mapping.xlsx")
        od_df.columns = od_df.columns.str.strip()
        
        # st.write("Yaha tak hua hai")
        
        # Helper: normalized column keys for fuzzy lookup
        def normalize(col_name: str) -> str:
            s = str(col_name).lower()
            s = re.sub(r"[^a-z0-9]", "", s)
            s = re.sub(r"hubs", "hub", s)
            return s
        
        normalized_cols = {normalize(c): c for c in od_df.columns}
        
        # === 66 RULES ===
        rules_66_data = [
            ["Origin Branch", "Origin Air Hub 1", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Origin Air Hub 2", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Origin Air Hub 1", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Origin Air Hub 2", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 1", "EP/BP", "Non_DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 2", "EP/BP", "Non_DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 1", "ES", "Non_DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 2", "ES", "Non_DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Branch", "EP/BP", "Non_DG", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Branch", "Destination Branch", "ES", "Non_DG", "White", "Standard Air", "Direct", "Y"],
            ["Origin Branch", "Origin Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Branch", "Origin Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Branch", "Origin Surface Hub 1", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Origin Surface Hub 2", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Origin Surface Hub 1", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Origin Surface Hub 2", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Branch", "Destination Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 1", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 1", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 2", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 2", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 1", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 1", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 2", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 2", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"],
            ["Origin Surface Hub 2", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"],
            ["Origin Surface Hub 1", "Destination Branch", "EP/BP", "DG", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Surface Hub 2", "Destination Branch", "EP/BP", "DG", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Surface Hub 1", "Destination Branch", "ES", "DG", "White", "Standard Air", "Direct", "Y"],
            ["Origin Surface Hub 2", "Destination Branch", "ES", "DG", "White", "Standard Air", "Direct", "Y"],
            ["Origin Air Hub 1", "Destination Air Hub 1", "EP/BP", "Non_DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 1", "EP/BP", "Non_DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Air Hub 2", "EP/BP", "Non_DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 2", "EP/BP", "Non_DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Air Hub 1", "ES", "Non_DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 1", "ES", "Non_DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Air Hub 2", "ES", "Non_DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 2", "ES", "Non_DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Surface Hub 1", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Surface Hub 1", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Surface Hub 2", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Surface Hub 2", "EP/BP", "DG", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Surface Hub 1", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Surface Hub 1", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Surface Hub 2", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Surface Hub 2", "ES", "DG", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Branch", "EP/BP", "Non_DG", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Air Hub 2", "Destination Branch", "EP/BP", "Non_DG", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Air Hub 1", "Destination Branch", "ES", "Non_DG", "White", "Standard Air", "Direct", "Y"],
            ["Origin Air Hub 2", "Destination Branch", "ES", "Non_DG", "White", "Standard Air", "Direct", "Y"],
            ["Destination Air Hub 1", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Air Hub 2", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Air Hub 1", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Air Hub 2", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Surface Hub 1", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Surface Hub 2", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Surface Hub 1", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Surface Hub 2", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Surface Hub 1", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"],
            ["Destination Surface Hub 2", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"]
        ]
        
        # === 42 RULES (for Surface mode only) ===
        rules_42_data = [
            ["Origin Branch", "Origin Air Hub 1", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Origin Air Hub 2", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Origin Air Hub 1", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Origin Air Hub 2", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 1", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 2", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 1", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Air Hub 2", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Branch", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Branch", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Origin Branch", "Origin Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Branch", "Origin Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Branch", "Destination Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Branch", "Destination Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 1", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 2", "Destination Surface Hub 2", "GP/BS", "Any", "White", "Surface", "Mixed", "N"],
            ["Origin Surface Hub 1", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"],
            ["Origin Surface Hub 2", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"],
            ["Origin Air Hub 1", "Destination Air Hub 1", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 1", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Air Hub 2", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 2", "EP/BP", "Any", "Red", "Premium Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Air Hub 1", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 1", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Air Hub 2", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 2", "Destination Air Hub 2", "ES", "Any", "White", "Standard Air", "Mixed", "N"],
            ["Origin Air Hub 1", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Air Hub 2", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Origin Air Hub 1", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Origin Air Hub 2", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Air Hub 1", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Air Hub 2", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Air Hub 1", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Air Hub 2", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Surface Hub 1", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Surface Hub 2", "Destination Branch", "EP/BP", "Any", "Red", "Premium Air", "Direct", "Y"],
            ["Destination Surface Hub 1", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Surface Hub 2", "Destination Branch", "ES", "Any", "White", "Standard Air", "Direct", "Y"],
            ["Destination Surface Hub 1", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"],
            ["Destination Surface Hub 2", "Destination Branch", "GP/BS", "Any", "White", "Surface", "Direct", "Y"]
        ]
        
        rules_cols = [
            "Bag Origin Type", "Bag Destination Type", "Product", "Goods Type",
            "Bag color", "CN Bkg Mode", "Bag Type", "Is_Direct"
        ]
        
        # === map rule label to column ===
        def find_column_for_rule_label(label, normalized_cols):
            nlabel = normalize(label)
            if nlabel in normalized_cols:
                return normalized_cols[nlabel]
            candidates = [orig_col for nc, orig_col in normalized_cols.items() if nlabel in nc or nc in nlabel]
            if len(candidates) == 1:
                return candidates[0]
            tokens = re.findall(r"[a-z]+", nlabel)
            for orig_col in od_df.columns:
                col_norm = normalize(orig_col)
                if all(tok in col_norm for tok in tokens[:min(2, len(tokens))]):
                    return orig_col
            return None
        
        # -----------------------------
        # Expand rules across OD rows
        # -----------------------------
        records = []
        missing_mappings = set()
        
        
        
        for idx, od in od_df.iterrows():
            origin = od.get("Origin Branch")
            dest = od.get("Destination Branch")
            mode = str(od.get("Possible Mode")).strip().capitalize() if pd.notna(od.get("Possible Mode")) else "Air"
        
            if pd.isna(origin) or pd.isna(dest):
                continue
        
            if mode == "Surface":
                rules_df = pd.DataFrame(rules_42_data, columns=rules_cols)
            else:
                rules_df = pd.DataFrame(rules_66_data, columns=rules_cols)
        
            for _, rule in rules_df.iterrows():
                bot = rule["Bag Origin Type"]
                bdt = rule["Bag Destination Type"]
        
                col_bot = find_column_for_rule_label(bot, normalized_cols)
                col_bdt = find_column_for_rule_label(bdt, normalized_cols)
        
                if col_bot is None or col_bdt is None:
                    missing_mappings.add((bot, bdt))
                    continue
        
                bag_origin = od.get(col_bot)
                bag_dest = od.get(col_bdt)
        
                if pd.isna(bag_origin) or pd.isna(bag_dest):
                    continue
        
                if origin == dest or bag_origin == bag_dest:
                    continue
        
                records.append({
                    "CN Origin": origin,
                    "CN Destination": dest,
                    "Bag Origin": bag_origin,
                    "is_direct": rule["Is_Direct"],
                    "packet/bag destination": bag_dest,
                    "Is_dg": rule["Goods Type"],
                    "Product": rule["Product"],
                    "Goods Type": rule["Goods Type"],
                    "Bag color": rule["Bag color"],
                    "CN Bkg Mode": rule["CN Bkg Mode"],
                    "Bag Type": rule["Bag Type"],
                    "Bag Origin Type": bot,
                    "Bag Destination Type": bdt,
                    "Possible Mode": mode
                })
        
        # report missing mappings
        #if missing_mappings:
            #st.write("⚠️ Some rule-labels couldn't be mapped to columns in the input file (samples):")
            #for i, mm in enumerate(list(missing_mappings)[:10], 1):
                
                #print(f"  {i}. {mm}")
            #print("Check OD_Hub_Mapping.xlsx headers for typos or mismatched names.")
        
        # Final dataframe
        final_df = pd.DataFrame(records)
        final_df = final_df[
            (final_df["CN Origin"] != final_df["CN Destination"]) &
            (final_df["Bag Origin"] != final_df["packet/bag destination"])
        ].reset_index(drop=True)
        
        final_df.to_csv("Final_Manifest_Rules_2_with mode.csv", index=False)
        st.write(f"Finished. Final manifest created & saved!")
        st.write(f"Total rows: {len(final_df):,}")
        
        
        # Define the path to the input CSV file
        #input_csv_path = "/content/drive/MyDrive/Manifest Lock/Final_Manifest_Rules_2_with mode.csv"
        
        # Read the CSV file into a DataFrame
        try:
            final_df_from_csv = final_df.copy()
            #print(f"Successfully loaded data from: {input_csv_path}")
            #print(f"Shape of loaded data: {final_df_from_csv.shape}")
        
            # Define the columns to group by
            group_cols = ["CN Origin", "CN Destination", "Bag Origin", "packet/bag destination", "Product", "Goods Type"]
        
            # Define the aggregation function
            def join_unique(x):
                return "/".join(x.dropna().astype(str).unique()) # Ensure aggregation works on strings
        
            # Group the DataFrame and apply the aggregation
            # Ensure grouping columns exist in the loaded DataFrame
            if all(col in final_df_from_csv.columns for col in group_cols):
                concatenated_df1 = final_df_from_csv.groupby(group_cols).agg(join_unique).reset_index()
        
                #output_path_concatenated1 = "/content/drive/MyDrive/Manifest Lock/Concatenated_Manifest_Rules_from_CSV.csv" # Changed output path
                concatenated_df1.to_csv("Concatenated_Manifest_Rules.csv", index=False)
                st.write(f"Concatenated DataFrame saved!")
                st.write(f"The length of the concatenated DataFrame is: {len(concatenated_df1)}")
        
            else:
                st.write(f"Error: One or more grouping columns {group_cols} not found in the loaded CSV file.")
        
        except FileNotFoundError:
            st.write(f"Error: Input CSV file not found at path. Please ensure the file exists.")
        except Exception as e:
            st.write(f"An error occurred during processing: {e}")
elif hub_mapping_flag == "No":
    st.write("Yaha Aaya!")

