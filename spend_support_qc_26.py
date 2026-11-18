import pandas as pd

from pathlib import Path

from openpyxl import load_workbook

from openpyxl.styles import Font, PatternFill, Alignment

from openpyxl.utils import get_column_letter

# ==================================================================
# SETTINGS

# ==================================================================
TRANSFORMATION_WEEKS = 26
EXCLUDED_DESCRIPTIONS = {
    "ACTUAL",
    "SEASONALITY",
    "PREDICTED",
    "INTERCEPT"
}
OUTPUT_COLUMNS = [
    "MODEL_KEY",
    "MODEL_KPI",
    "GEOGRAPHY",
    "CUSTOMER_SEGMENT",
    "INTERACTION_CHANNEL",
    "VARIABLE",
    "WEEK_END",
    "SPEND",
    "SUPPORT",
    "VOLUME",
    "TIMESTAMP_UTC",
    "LOB_METRIC"
]

# ==================================================================
# 1. GET FILE PATHS

# ==================================================================
ads_path_input = input(
    "\nEnter the full path of the ADS file: "
).strip().strip('"')
mapping_path_input = input(
    "Enter the full path of the Variable Mapping file: "
).strip().strip('"')
flat_file_path_input = input(
    "Enter the full path of the Output Spend/Support Flat File: "
).strip().strip('"')
ads_path = Path(ads_path_input)
mapping_path = Path(mapping_path_input)
flat_file_path = Path(flat_file_path_input)

# ==================================================================
# 2. CHECK FILES

# ==================================================================
if not ads_path.exists():
    raise FileNotFoundError(
        f"\nADS file not found:\n{ads_path}"
    )
if not mapping_path.exists():
    raise FileNotFoundError(
        f"\nVariable Mapping file not found:\n{mapping_path}"
    )
if not flat_file_path.exists():
    raise FileNotFoundError(
        f"\nOutput Spend/Support Flat File not found:\n{flat_file_path}"
    )

# ==================================================================
# 2A. READ GENERATED FLAT FILE

# ==================================================================
print("\nReading generated Spend/Support flat file...")
flat_file = pd.read_csv(
    flat_file_path,
    low_memory=False
)
print(
    f"Flat file rows : {len(flat_file):,}"
)
# The production file merges multiple ADS variables when they have
# the same ADI Level 3, but only within the same Spend/Support
# MODEL_KEY.  These are the exact columns used by production for
# the final row grain.
FLAT_FILE_GROUP_COLUMNS = [
    "MODEL_KEY",
    "MODEL_KPI",
    "GEOGRAPHY",
    "CUSTOMER_SEGMENT",
    "INTERACTION_CHANNEL",
    "VARIABLE",
    "WEEK_END",
    "TIMESTAMP_UTC",
    "LOB_METRIC"
]
required_flat_columns = (
    FLAT_FILE_GROUP_COLUMNS
    + ["SPEND", "SUPPORT", "VOLUME"]
)
missing_flat_columns = [
    col
    for col in required_flat_columns
    if col not in flat_file.columns
]
if missing_flat_columns:
    raise ValueError(
        "\nGenerated flat file is missing required columns:\n"
        + "\n".join(
            f" - {col}"
            for col in missing_flat_columns
        )
    )
# Convert numeric fields exactly as they are treated in production.
flat_file["SPEND"] = pd.to_numeric(
    flat_file["SPEND"],
    errors="coerce"
).fillna(0)
flat_file["SUPPORT"] = pd.to_numeric(
    flat_file["SUPPORT"],
    errors="coerce"
).fillna(0)
flat_file["VOLUME"] = pd.to_numeric(
    flat_file["VOLUME"],
    errors="coerce"
).fillna(0)
# This is the actual post-ADI-merge row count.
flat_file_row_count = len(flat_file)
flat_file_total_spend = flat_file["SPEND"].sum()
flat_file_total_support = flat_file["SUPPORT"].sum()
flat_file_structure_check = (
    "PASS"
    if list(flat_file.columns) == OUTPUT_COLUMNS
    else "FAIL"
)
flat_file_volume_zero_check = (
    "PASS"
    if (flat_file["VOLUME"] == 0).all()
    else "FAIL"
)
print(
    f"Post-ADI-merge rows : {flat_file_row_count:,}"
)
print(
    f"Flat file Spend     : {flat_file_total_spend:,.2f}"
)
print(
    f"Flat file Support   : {flat_file_total_support:,.2f}"
)

# ==================================================================
# 3. GET LOB FROM ADS FILE NAME ONLY

# ==================================================================
# Example:
#
# USAA_P&C_Master_ADS.csv
#
# USAA | P&C | Master | ADS
#        ^
#        LOB
#
# IMPORTANT:
# Only the file name is used.
# The folder path is NOT used.
file_name = ads_path.name
file_stem = ads_path.stem
file_parts = file_stem.split("_")
if len(file_parts) < 2:
    raise ValueError(
        f"\nUnable to determine LOB from ADS file name:\n"
        f"{file_name}\n\n"
        f"Expected format similar to:\n"
        f"USAA_P&C_Master_ADS.csv"
    )
lob = file_parts[1].strip()
if not lob:
    raise ValueError(
        f"\nLOB could not be determined from:\n"
        f"{file_name}"
    )

# ==================================================================
# 4. AUTOMATIC MODEL KEYS

# ==================================================================
support_model_key = f"{lob} - Support"
spend_model_key = f"{lob} - Spend"
brand_support_model_key = "Brand - Support"
brand_spend_model_key = "Brand - Spend"

# ==================================================================
# 5. QC OUTPUT PATH

# ==================================================================
# QC file is saved in the same folder as the ADS file.
qc_output = (
    ads_path.parent /
    f"USAA_{lob}_Spend_Support_Flat_File_QC.xlsx"
)

# ==================================================================
# 6. DISPLAY INFORMATION

# ==================================================================
print("\n" + "=" * 70)
print("SPEND / SUPPORT QC")
print("=" * 70)
print(f"ADS File          : {ads_path}")
print(f"Variable Mapping  : {mapping_path}")
print(f"Output Flat File  : {flat_file_path}")
print(f"Detected LOB      : {lob}")
print(f"Support Model Key : {support_model_key}")
print(f"Spend Model Key   : {spend_model_key}")
print(f"QC Output         : {qc_output}")

# ==================================================================
# 7. READ ADS FILE

# ==================================================================
print("\n" + "=" * 70)
print("READING ADS FILE")
print("=" * 70)
ads = pd.read_csv(
    ads_path,
    low_memory=False
)
print(
    f"ADS rows    : {len(ads):,}"
)
print(
    f"ADS columns : {len(ads.columns):,}"
)

# ==================================================================
# 8. IDENTIFY ADS COLUMNS

# ==================================================================
# Standard ADS structure:
#
# A = Period
# B = Geography
# C onwards = ADS variables
if len(ads.columns) < 3:
    raise ValueError(
        "ADS file must contain at least 3 columns."
    )
period_col = ads.columns[0]
geo_col = ads.columns[1]
ads_variable_columns = list(
    ads.columns[2:]
)
print("\n" + "=" * 70)
print("ADS COLUMN STRUCTURE")
print("=" * 70)
print(
    f"Period column       : {period_col}"
)
print(
    f"Geography column    : {geo_col}"
)
print(
    f"First ADS variable  : "
    f"{ads_variable_columns[0]}"
)
print(
    f"Last ADS variable   : "
    f"{ads_variable_columns[-1]}"
)
print(
    f"Total ADS variables : "
    f"{len(ads_variable_columns):,}"
)

# ==================================================================
# 9. CLEAN PERIOD COLUMN

# ==================================================================
print("\nCleaning dates...")
ads[period_col] = pd.to_datetime(
    ads[period_col],
    errors="coerce"
)
invalid_dates = ads[period_col].isna().sum()
if invalid_dates > 0:
    print(
        f"WARNING: {invalid_dates:,} invalid dates found."
    )
ads = ads.dropna(
    subset=[period_col]
).copy()

# ==================================================================
# 10. CLEAN GEOGRAPHY

# ==================================================================
ads[geo_col] = (
    ads[geo_col]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ==================================================================
# 11. HANDLE TOTAL US

# ==================================================================
#
# If the ADS contains other geographies in addition to
# 00_00_Total_US:
#     -> Remove 00_00_Total_US
#
# If 00_00_Total_US is the ONLY geography:
#     -> Keep 00_00_Total_US
#
# This allows Total US to be retained for LOBs where it is the
# only geography available.

# ==================================================================
unique_geographies = (
    ads[geo_col]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)
non_total_us_geographies = [
    geo
    for geo in unique_geographies
    if geo != "00_00_Total_US"
]
initial_rows = len(ads)
if len(non_total_us_geographies) > 0:
    ads = ads[
        ads[geo_col] != "00_00_Total_US"
    ].copy()
    removed_total_us = (
        initial_rows - len(ads)
    )
    print(
        f"Other geographies found : "
        f"{len(non_total_us_geographies):,}"
    )
    print(
        f"Removed Total US rows   : "
        f"{removed_total_us:,}"
    )
else:
    removed_total_us = 0
    print(
        "Only 00_00_Total_US geography found."
    )
    print(
        "Total US will be retained."
    )

# ==================================================================
# 12. CLEAN GEOGRAPHY

# ==================================================================
# Example:
#
# 01_33_State_North_Carolina
#             ↓
# North_Carolina
#
# 01_30_State_New_Jersey
#             ↓
# New_Jersey
def clean_geography(geo):
    geo = str(geo).strip()
    if geo == "00_00_Total_US":
        return "Total_US"
    if "_State_" in geo:
        return geo.split(
            "_State_",
            1
        )[1]
    return geo
ads[geo_col] = ads[
    geo_col
].apply(
    clean_geography
)
# Keep a complete copy of the ADS for full-period QC totals.
# All-zero checks and full-period totals must use the complete ADS,
# before the 26-week transformation period is removed.
full_ads = ads.copy()

# ==================================================================
# 13. ADS DATE RANGE

# ==================================================================
if ads.empty:
    raise ValueError(
        "No ADS data remains after Total US handling."
    )
ads_start_date = ads[
    period_col
].min()
ads_end_date = ads[
    period_col
].max()
print(
    f"ADS Start Date : "
    f"{ads_start_date.strftime('%m/%d/%Y')}"
)
print(
    f"ADS End Date   : "
    f"{ads_end_date.strftime('%m/%d/%Y')}"
)

# ==================================================================
# 14. READ VARIABLE MAPPING

# ==================================================================
print("\n" + "=" * 70)
print("READING VARIABLE MAPPING")
print("=" * 70)
mapping_variable = pd.read_excel(
    mapping_path,
    sheet_name="Variable"
)
mapping_model_key = pd.read_excel(
    mapping_path,
    sheet_name="Model Key"
)
print(
    f"Variable tab rows  : "
    f"{len(mapping_variable):,}"
)
print(
    f"Model Key tab rows : "
    f"{len(mapping_model_key):,}"
)

# ==================================================================
# 15. CHECK VARIABLE MAPPING

# ==================================================================
if len(mapping_variable.columns) < 4:
    raise ValueError(
        "The 'Variable' sheet must contain "
        "at least 4 columns."
    )
# Column A = Support 3.3
# Column B = New Spend 3.3
# Column C = ADI Level 3
# Column D = Customer Segment
mapping_support_col = (
    mapping_variable.columns[0]
)
mapping_spend_col = (
    mapping_variable.columns[1]
)
mapping_adi_col = (
    mapping_variable.columns[2]
)
mapping_segment_col = (
    mapping_variable.columns[3]
)

# ==================================================================
# 16. CLEAN VARIABLE MAPPING

# ==================================================================
mapping_variable[
    mapping_support_col
] = (
    mapping_variable[
        mapping_support_col
    ]
    .fillna("")
    .astype(str)
    .str.strip()
)
mapping_variable[
    mapping_spend_col
] = (
    mapping_variable[
        mapping_spend_col
    ]
    .fillna("")
    .astype(str)
    .str.strip()
)
mapping_variable[
    mapping_adi_col
] = (
    mapping_variable[
        mapping_adi_col
    ]
    .fillna("")
    .astype(str)
    .str.strip()
)
mapping_variable[
    mapping_segment_col
] = (
    mapping_variable[
        mapping_segment_col
    ]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ==================================================================
# 17. CREATE MAPPING LOOKUPS

# ==================================================================
support_lookup = {}
spend_lookup = {}
for _, row in mapping_variable.iterrows():
    support_name = str(
        row[mapping_support_col]
    ).strip()
    spend_name = str(
        row[mapping_spend_col]
    ).strip()
    adi_level_3 = str(
        row[mapping_adi_col]
    ).strip()
    customer_segment = str(
        row[mapping_segment_col]
    ).strip()
    if (
        not customer_segment
        or customer_segment.lower() == "nan"
    ):
        customer_segment = "n/a"
    if support_name:
        support_lookup[
            support_name
        ] = {
            "adi_level_3": adi_level_3,
            "customer_segment": customer_segment
        }
    if spend_name:
        spend_lookup[
            spend_name
        ] = {
            "adi_level_3": adi_level_3,
            "customer_segment": customer_segment
        }
print(
    f"Support mappings : "
    f"{len(support_lookup):,}"
)
print(
    f"Spend mappings   : "
    f"{len(spend_lookup):,}"
)

# ==================================================================
# 18. MODEL KEY MAPPING

# ==================================================================
if len(mapping_model_key.columns) < 5:
    raise ValueError(
        "The 'Model Key' sheet must contain at least 5 columns."
    )
model_key_mapping = mapping_model_key.iloc[
    :,
    [0, 2, 3, 4]
].copy()
model_key_mapping.columns = [
    "MODEL_KEY",
    "MODEL_KPI",
    "CUSTOMER_SEGMENT",
    "INTERACTION_CHANNEL"
]
for col in model_key_mapping.columns:
    model_key_mapping[col] = (
        model_key_mapping[col]
        .fillna("")
        .astype(str)
        .str.strip()
    )
model_key_mapping = (
    model_key_mapping
    .drop_duplicates(
        subset=["MODEL_KEY"],
        keep="first"
    )
)
model_key_lookup = {}
for _, row in model_key_mapping.iterrows():
    model_key_value = row["MODEL_KEY"]
    if not model_key_value:
        continue
    model_key_lookup[model_key_value] = {
        "MODEL_KPI": (
            row["MODEL_KPI"]
            if row["MODEL_KPI"]
            else "n/a"
        ),
        "CUSTOMER_SEGMENT": (
            row["CUSTOMER_SEGMENT"]
            if row["CUSTOMER_SEGMENT"]
            else "n/a"
        ),
        "INTERACTION_CHANNEL": (
            row["INTERACTION_CHANNEL"]
            if row["INTERACTION_CHANNEL"]
            else "n/a"
        )
    }
print(
    f"Model Key mappings : "
    f"{len(model_key_lookup):,}"
)

# ==================================================================

# ==================================================================
print("\n" + "=" * 70)
print("CHECKING ALL-ZERO VARIABLES")
print("=" * 70)
all_zero_variables = []
total_ads_variables = len(
    ads_variable_columns
)
for counter, variable in enumerate(
    ads_variable_columns,
    start=1
):
    values = pd.to_numeric(
        ads[variable],
        errors="coerce"
    ).fillna(0)
    if (values == 0).all():
        all_zero_variables.append(
            variable
        )
    if (
        counter % 100 == 0
        or counter == total_ads_variables
    ):
        print(
            f"Checked "
            f"{counter:,}/"
            f"{total_ads_variables:,}"
        )
print(
    f"\nAll-zero variables : "
    f"{len(all_zero_variables):,}"
)

# ==================================================================
# 19. SKIP TRANSFORMATION PERIOD

# ==================================================================
# The first 26 unique weekly periods are the transformation period.
# This is determined dynamically from the ADS dates; no fixed start
# date is used. The all-zero check above has already been completed
# against the full ADS.
unique_weeks = (
    full_ads[period_col]
    .drop_duplicates()
    .sort_values()
    .tolist()
)
if len(unique_weeks) <= TRANSFORMATION_WEEKS:
    raise ValueError(
        f"ADS contains only {len(unique_weeks):,} unique weeks. "
        f"At least {TRANSFORMATION_WEEKS + 1:,} weeks are required "
        f"to skip the first {TRANSFORMATION_WEEKS} transformation weeks."
    )
transformation_end_date = unique_weeks[TRANSFORMATION_WEEKS - 1]
ads = full_ads[
    full_ads[period_col] > transformation_end_date
].copy()
remaining_weeks = (
    ads[period_col]
    .drop_duplicates()
    .nunique()
)
print(
    f"\nTransformation weeks skipped : {TRANSFORMATION_WEEKS:,}"
)
print(
    f"Transformation period end    : "
    f"{transformation_end_date.strftime('%m/%d/%Y')}"
)
print(
    f"Weeks remaining for QC      : {remaining_weeks:,}"
)

# ==================================================================
# 20. MATCH VARIABLES TO MAPPING

# ==================================================================
print("\n" + "=" * 70)
print("MATCHING ADS VARIABLES TO MAPPING")
print("=" * 70)
selected_variables = []
excluded_variables = []
not_in_mapping_variables = []
all_zero_set = set(
    all_zero_variables
)
for variable in ads_variable_columns:
    is_brand = "_BRD_" in variable
    # --------------------------------------------------------------
    # SUPPORT
    # --------------------------------------------------------------
    if variable in support_lookup:
        mapping_info = support_lookup[
            variable
        ]
        adi_level_3 = mapping_info[
            "adi_level_3"
        ]
        customer_segment = mapping_info[
            "customer_segment"
        ]
        if (
            adi_level_3.strip().upper()
            in EXCLUDED_DESCRIPTIONS
        ):
            excluded_variables.append({
                "ADS_VARIABLE": variable,
                "TYPE": "Support",
                "ADI_LEVEL_3": adi_level_3,
                "CUSTOMER_SEGMENT": customer_segment,
                "REASON": (
                    "ADI Level 3 is excluded "
                    "(Actual/Seasonality/"
                    "Predicted/Intercept)"
                )
            })
            continue
        if variable in all_zero_set:
            continue
        selected_variables.append({
            "ads_variable": variable,
            "model_key": (
                brand_support_model_key
                if is_brand
                else support_model_key
            ),
            "adi_level_3": adi_level_3,
            "customer_segment": customer_segment,
            "type": "Support",
            "is_brand": is_brand
        })
        continue
    # --------------------------------------------------------------
    # SPEND
    # --------------------------------------------------------------
    if variable in spend_lookup:
        mapping_info = spend_lookup[
            variable
        ]
        adi_level_3 = mapping_info[
            "adi_level_3"
        ]
        customer_segment = mapping_info[
            "customer_segment"
        ]
        if (
            adi_level_3.strip().upper()
            in EXCLUDED_DESCRIPTIONS
        ):
            excluded_variables.append({
                "ADS_VARIABLE": variable,
                "TYPE": "Spend",
                "ADI_LEVEL_3": adi_level_3,
                "CUSTOMER_SEGMENT": customer_segment,
                "REASON": (
                    "ADI Level 3 is excluded "
                    "(Actual/Seasonality/"
                    "Predicted/Intercept)"
                )
            })
            continue
        if variable in all_zero_set:
            continue
        selected_variables.append({
            "ads_variable": variable,
            "model_key": (
                brand_spend_model_key
                if is_brand
                else spend_model_key
            ),
            "adi_level_3": adi_level_3,
            "customer_segment": customer_segment,
            "type": "Spend",
            "is_brand": is_brand
        })
        continue
    # --------------------------------------------------------------
    # NOT IN MAPPING
    # --------------------------------------------------------------
    not_in_mapping_variables.append(
        variable
    )
print(
    f"\nSelected variables : "
    f"{len(selected_variables):,}"
)
print(
    f"Excluded variables : "
    f"{len(excluded_variables):,}"
)
print(
    f"Not in mapping     : "
    f"{len(not_in_mapping_variables):,}"
)

# ==================================================================
# 21. PREPARE QC STORAGE

# ==================================================================

# ==================================================================
variable_qc_rows = []
geo_qc_rows = []
zero_spend_support_rows = []

# ==================================================================
# 22. RUN QC CALCULATIONS

# ==================================================================
print("\n" + "=" * 70)
print("RUNNING QC CALCULATIONS")
print("=" * 70)
total_variables = len(
    selected_variables
)
for counter, item in enumerate(
    selected_variables,
    start=1
):
    variable = item[
        "ads_variable"
    ]
    model_key = item[
        "model_key"
    ]
    adi_level_3 = item[
        "adi_level_3"
    ]
    customer_segment = item[
        "customer_segment"
    ]
    variable_type = item[
        "type"
    ]
    model_info = model_key_lookup.get(
        model_key,
        {}
    )
    model_kpi = model_info.get(
        "MODEL_KPI",
        "n/a"
    )
    mapped_customer_segment = model_info.get(
        "CUSTOMER_SEGMENT",
        customer_segment
    )
    interaction_channel = model_info.get(
        "INTERACTION_CHANNEL",
        "n/a"
    )
    if not mapped_customer_segment:
        mapped_customer_segment = customer_segment
    is_brand = item.get(
        "is_brand",
        False
    )
    lob_metric = (
        "Brand"
        if is_brand
        else f"{lob} Spend-Support"
    )
    # --------------------------------------------------------------
    # Create temporary data
    #
    # This is only used for QC.
    # It is NOT saved as a flat file.
    # --------------------------------------------------------------
    temp = ads[
        [
            period_col,
            geo_col,
            variable
        ]
    ].copy()
    temp.rename(
        columns={
            period_col: "WEEK_END",
            geo_col: "GEOGRAPHY",
            variable: "VALUE"
        },
        inplace=True
    )
    # --------------------------------------------------------------
    # Numeric conversion
    # --------------------------------------------------------------
    temp["VALUE"] = pd.to_numeric(
        temp["VALUE"],
        errors="coerce"
    ).fillna(0)
    # --------------------------------------------------------------
    # QC fields
    # --------------------------------------------------------------
    temp["MODEL_KEY"] = (
        model_key
    )
    temp["MODEL_KPI"] = (
        model_kpi
    )
    temp["VARIABLE"] = (
        adi_level_3
    )
    temp["CUSTOMER_SEGMENT"] = (
        mapped_customer_segment
    )
    temp["INTERACTION_CHANNEL"] = (
        interaction_channel
    )
    temp["ADSTOCK"] = "n/a"
    temp["POWER"] = "n/a"
    temp["LAG"] = "n/a"
    temp["ROLLING_AVERAGE"] = "n/a"
    temp["VOLUME"] = 0
    temp["LOB_METRIC"] = (
        lob_metric
    )
    # --------------------------------------------------------------
    # Spend / Support
    # --------------------------------------------------------------
    if variable_type == "Spend":
        temp["SPEND"] = (
            temp["VALUE"]
        )
        temp["SUPPORT"] = 0
    else:
        temp["SPEND"] = 0
        temp["SUPPORT"] = (
            temp["VALUE"]
        )
    # --------------------------------------------------------------
    # Dates
    # --------------------------------------------------------------
    temp["WEEK_END"] = pd.to_datetime(
        temp["WEEK_END"],
        errors="coerce"
    )
    # ==============================================================
    # VARIABLE QC
    # ==============================================================
    variable_total_spend = (
        temp["SPEND"].sum()
    )
    variable_total_support = (
        temp["SUPPORT"].sum()
    )
    variable_start_date = (
        temp["WEEK_END"].min()
    )
    variable_end_date = (
        temp["WEEK_END"].max()
    )
    # --------------------------------------------------------------
    # Full-period totals
    # --------------------------------------------------------------
    full_temp = full_ads[
        [
            period_col,
            geo_col,
            variable
        ]
    ].copy()
    full_temp.rename(
        columns={
            period_col: "WEEK_END",
            geo_col: "GEOGRAPHY",
            variable: "VALUE"
        },
        inplace=True
    )
    full_temp["VALUE"] = pd.to_numeric(
        full_temp["VALUE"],
        errors="coerce"
    ).fillna(0)
    if variable_type == "Spend":
        full_temp["SPEND"] = full_temp["VALUE"]
        full_temp["SUPPORT"] = 0
    else:
        full_temp["SPEND"] = 0
        full_temp["SUPPORT"] = full_temp["VALUE"]
    full_temp["WEEK_END"] = pd.to_datetime(
        full_temp["WEEK_END"],
        errors="coerce"
    )
    full_total_spend = full_temp["SPEND"].sum()
    full_total_support = full_temp["SUPPORT"].sum()
    full_start_date = full_temp["WEEK_END"].min()
    full_end_date = full_temp["WEEK_END"].max()
    variable_qc_rows.append({
        "MODEL_KEY": model_key,
        "MODEL_KPI": model_kpi,
        "VARIABLE": adi_level_3,
        "ADS_VARIABLE": variable,
        "TYPE": variable_type,
        "CUSTOMER_SEGMENT":
            mapped_customer_segment,
        "INTERACTION_CHANNEL":
            interaction_channel,
        "LOB_METRIC":
            lob_metric,
        "START_DATE":
            variable_start_date.strftime(
                "%m/%d/%Y"
            ),
        "END_DATE":
            variable_end_date.strftime(
                "%m/%d/%Y"
            ),
        "TOTAL_ROWS":
            len(temp),
        "TOTAL_SPEND":
            variable_total_spend,
        "TOTAL_SUPPORT":
            variable_total_support,
        "FULL_PERIOD_START_DATE":
            full_start_date.strftime(
                "%m/%d/%Y"
            ),
        "FULL_PERIOD_END_DATE":
            full_end_date.strftime(
                "%m/%d/%Y"
            ),
        "FULL_PERIOD_TOTAL_ROWS":
            len(full_temp),
        "FULL_PERIOD_TOTAL_SPEND":
            full_total_spend,
        "FULL_PERIOD_TOTAL_SUPPORT":
            full_total_support
    })
    # ==============================================================
    # GEO QC
    # ==============================================================
    geo_grouped = (
        temp
        .groupby(
            "GEOGRAPHY",
            dropna=False
        )
        .agg(
            TOTAL_SPEND=(
                "SPEND",
                "sum"
            ),
            TOTAL_SUPPORT=(
                "SUPPORT",
                "sum"
            ),
            ROW_COUNT=(
                "GEOGRAPHY",
                "size"
            )
        )
        .reset_index()
    )
    for _, geo_row in (
        geo_grouped.iterrows()
    ):
        geo_name = geo_row[
            "GEOGRAPHY"
        ]
        geo_data = temp[
            temp["GEOGRAPHY"] == geo_name
        ].copy()
        full_geo_data = full_temp[
            full_temp["GEOGRAPHY"] == geo_name
        ].copy()
        full_geo_spend = full_geo_data["SPEND"].sum()
        full_geo_support = full_geo_data["SUPPORT"].sum()
        full_geo_start_date = full_geo_data["WEEK_END"].min()
        full_geo_end_date = full_geo_data["WEEK_END"].max()
        geo_qc_rows.append({
            "MODEL_KEY": model_key,
            "MODEL_KPI": model_kpi,
            "VARIABLE": adi_level_3,
            "ADS_VARIABLE": variable,
            "TYPE": variable_type,
            "GEOGRAPHY": geo_name,
            "CUSTOMER_SEGMENT":
                mapped_customer_segment,
            "INTERACTION_CHANNEL":
                interaction_channel,
            "LOB_METRIC":
                lob_metric,
            "TOTAL_SPEND":
                geo_row["TOTAL_SPEND"],
            "TOTAL_SUPPORT":
                geo_row["TOTAL_SUPPORT"],
            "ROW_COUNT":
                geo_row["ROW_COUNT"],
            "FULL_PERIOD_START_DATE":
                full_geo_start_date.strftime(
                    "%m/%d/%Y"
                ),
            "FULL_PERIOD_END_DATE":
                full_geo_end_date.strftime(
                    "%m/%d/%Y"
                ),
            "FULL_PERIOD_TOTAL_ROWS":
                len(full_geo_data),
            "FULL_PERIOD_TOTAL_SPEND":
                full_geo_spend,
            "FULL_PERIOD_TOTAL_SUPPORT":
                full_geo_support
        })
    # ==============================================================
    # ZERO SPEND / SUPPORT
    # ==============================================================
    if (
        variable_total_spend == 0
        and variable_total_support == 0
    ):
        zero_spend_support_rows.append({
            "MODEL_KEY": model_key,
            "VARIABLE": adi_level_3,
            "ADS_VARIABLE": variable,
            "TYPE": variable_type,
            "CUSTOMER_SEGMENT":
                mapped_customer_segment,
            "TOTAL_SPEND":
                variable_total_spend,
            "TOTAL_SUPPORT":
                variable_total_support
        })
    # ==============================================================
    # PROGRESS
    # ==============================================================
    if (
        counter % 10 == 0
        or counter == total_variables
    ):
        percent = (
            counter /
            total_variables *
            100
        )
        print(
            f"Progress: "
            f"{counter:,}/"
            f"{total_variables:,} "
            f"({percent:.1f}%)"
        )

# ==================================================================
# 23. CREATE QC DATAFRAMES

# ==================================================================
variable_qc_df = pd.DataFrame(
    variable_qc_rows
)
geo_qc_df = pd.DataFrame(
    geo_qc_rows
)
zero_spend_support_df = pd.DataFrame(
    zero_spend_support_rows
)
excluded_variables_df = pd.DataFrame(
    excluded_variables
)
all_zero_variables_df = pd.DataFrame({
    "ADS_VARIABLE":
        all_zero_variables
})
not_in_mapping_df = pd.DataFrame({
    "ADS_VARIABLE":
        not_in_mapping_variables
})

# ==================================================================
# 24. QC SUMMARY

# ==================================================================
# IMPORTANT:
# Do NOT sum VARIABLE_QC TOTAL_SPEND / TOTAL_SUPPORT here.
# VARIABLE_QC is intentionally at ADS-variable level, so duplicate
# ADI Level 3 variables would be counted more than once.
#
# The production flat file has already merged duplicate ADI Level 3
# records. Therefore QC summary totals must come from the generated
# flat file.
total_qc_spend = flat_file_total_spend
total_qc_support = flat_file_total_support
# Post-merge row count from the actual production output.
post_merge_qc_rows = flat_file_row_count
summary_data = [
    {
        "CHECK": "LOB",
        "VALUE": lob
    },
    {
        "CHECK": "Brand Support Model Key",
        "VALUE": brand_support_model_key
    },
    {
        "CHECK": "Brand Spend Model Key",
        "VALUE": brand_spend_model_key
    },
    {
        "CHECK": "Support Model Key",
        "VALUE": support_model_key
    },
    {
        "CHECK": "Spend Model Key",
        "VALUE": spend_model_key
    },
    {
        "CHECK": "ADS File",
        "VALUE": file_name
    },
    {
        "CHECK": "Variable Mapping File",
        "VALUE": mapping_path.name
    },
    {
        "CHECK":
            "Total ADS Rows After Total US Handling",
        "VALUE":
            len(ads)
    },
    {
        "CHECK":
            "Total ADS Variables",
        "VALUE":
            len(ads_variable_columns)
    },
    {
        "CHECK":
            "Selected Variables",
        "VALUE":
            len(selected_variables)
    },
    {
        "CHECK":
            "Excluded Variables",
        "VALUE":
            len(excluded_variables)
    },
    {
        "CHECK":
            "All-Zero Variables",
        "VALUE":
            len(all_zero_variables)
    },
    {
        "CHECK":
            "Not in Mapping",
        "VALUE":
            len(not_in_mapping_variables)
    },
    {
        "CHECK":
            "QC Rows - Post ADI Merge",
        "VALUE":
            post_merge_qc_rows
    },
    {
        "CHECK":
            "Variable-Level QC Rows Before ADI Merge",
        "VALUE":
            variable_qc_df[
                "TOTAL_ROWS"
            ].sum()
    },
    {
        "CHECK":
            "Flat File Row Count Check",
        "VALUE":
            (
                "PASS"
                if post_merge_qc_rows == len(flat_file)
                else "FAIL"
            )
    },
    {
        "CHECK":
            "Total Spend",
        "VALUE":
            total_qc_spend
    },
    {
        "CHECK":
            "Total Support",
        "VALUE":
            total_qc_support
    },
    {
        "CHECK":
            "Flat File Structure",
        "VALUE":
            flat_file_structure_check
    },
    {
        "CHECK":
            "Flat File Volume = 0",
        "VALUE":
            flat_file_volume_zero_check
    },
    {
        "CHECK":
            "Flat File KPI Fields = n/a",
        "VALUE":
            "flat_file_kpi_fields_check"
    },
    {
        "CHECK":
            "Transformation Weeks Skipped",
        "VALUE":
            TRANSFORMATION_WEEKS
    },
    {
        "CHECK":
            "Transformation Period End",
        "VALUE":
            transformation_end_date.strftime(
                "%m/%d/%Y"
            )
    },
    {
        "CHECK":
            "QC Output Start Date",
        "VALUE":
            ads[period_col].min().strftime(
                "%m/%d/%Y"
            )
    },
    {
        "CHECK":
            "QC Output End Date",
        "VALUE":
            ads[period_col].max().strftime(
                "%m/%d/%Y"
            )
    },
    {
        "CHECK":
            "Full ADS Start Date",
        "VALUE":
            full_ads[period_col].min().strftime(
                "%m/%d/%Y"
            )
    },
    {
        "CHECK":
            "Full ADS End Date",
        "VALUE":
            full_ads[period_col].max().strftime(
                "%m/%d/%Y"
            )
    }
]
qc_summary_df = pd.DataFrame(
    summary_data
)

# ==================================================================
# 25. WRITE QC EXCEL ONLY

# ==================================================================
print("\n" + "=" * 70)
print("CREATING QC WORKBOOK")
print("=" * 70)
with pd.ExcelWriter(
    qc_output,
    engine="openpyxl"
) as writer:
    qc_summary_df.to_excel(
        writer,
        sheet_name="QC_SUMMARY",
        index=False
    )
    variable_qc_df.to_excel(
        writer,
        sheet_name="VARIABLE_QC",
        index=False
    )
    geo_qc_df.to_excel(
        writer,
        sheet_name="GEO_SPEND_SUPPORT",
        index=False
    )
    zero_spend_support_df.to_excel(
        writer,
        sheet_name="ZERO_SPEND_SUPPORT",
        index=False
    )
    excluded_variables_df.to_excel(
        writer,
        sheet_name="EXCLUDED_VARIABLES",
        index=False
    )
    all_zero_variables_df.to_excel(
        writer,
        sheet_name="ALL_ZERO_VARIABLES",
        index=False
    )
    not_in_mapping_df.to_excel(
        writer,
        sheet_name="NOT_IN_MAPPING",
        index=False
    )
    # --------------------------------------------------------------
    # POST-ADI-MERGE FLAT FILE SUMMARY
    # --------------------------------------------------------------
    flat_file_summary_df = pd.DataFrame([
        {
            "CHECK": "Generated Flat File",
            "VALUE": flat_file_path.name
        },
        {
            "CHECK": "Post-ADI-Merge Row Count",
            "VALUE": flat_file_row_count
        },
        {
            "CHECK": "Flat File Total Spend",
            "VALUE": flat_file_total_spend
        },
        {
            "CHECK": "Flat File Total Support",
            "VALUE": flat_file_total_support
        },
        {
            "CHECK": "Row Count Validation",
            "VALUE": (
                "PASS"
                if flat_file_row_count == len(flat_file)
                else "FAIL"
            )
        },
        {
            "CHECK": "Flat File Structure",
            "VALUE": flat_file_structure_check
        },
        {
            "CHECK": "Flat File Volume = 0",
            "VALUE": flat_file_volume_zero_check
        },
        {
            "CHECK": "Flat File KPI Fields = n/a",
            "VALUE": "flat_file_kpi_fields_check"
        }
    ])
    flat_file_summary_df.to_excel(
        writer,
        sheet_name="POST_ADI_MERGE_QC",
        index=False
    )

# ==================================================================
# 26. FORMAT QC WORKBOOK

# ==================================================================
print("Formatting QC workbook...")
wb = load_workbook(
    qc_output
)
header_fill = PatternFill(
    fill_type="solid",
    fgColor="4472C4"
)
header_font = Font(
    bold=True,
    color="FFFFFF"
)
header_alignment = Alignment(
    horizontal="center",
    vertical="center"
)
for ws in wb.worksheets:
    # --------------------------------------------------------------
    # Header formatting
    # --------------------------------------------------------------
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = (
            header_alignment
        )
    # --------------------------------------------------------------
    # Freeze top row
    # --------------------------------------------------------------
    ws.freeze_panes = "A2"
    # --------------------------------------------------------------
    # Autofilter
    # --------------------------------------------------------------
    if ws.max_row > 1:
        ws.auto_filter.ref = (
            ws.dimensions
        )
    # --------------------------------------------------------------
    # Header height
    # --------------------------------------------------------------
    ws.row_dimensions[1].height = 22
    # --------------------------------------------------------------
    # Auto column width
    # --------------------------------------------------------------
    for column_cells in ws.columns:
        max_length = 0
        column_letter = (
            get_column_letter(
                column_cells[0].column
            )
        )
        for cell in column_cells:
            try:
                cell_length = len(
                    str(cell.value)
                )
                if cell_length > max_length:
                    max_length = (
                        cell_length
                    )
            except Exception:
                pass
        adjusted_width = min(
            max(max_length + 2, 10),
            50
        )
        ws.column_dimensions[
            column_letter
        ].width = adjusted_width

# ==================================================================
# 27. SAVE QC FILE

# ==================================================================
wb.save(
    qc_output
)

# ==================================================================
# 28. FINAL MESSAGE

# ==================================================================
print("\n" + "=" * 70)
print("QC COMPLETED SUCCESSFULLY")
print("=" * 70)
print(
    f"\nDetected LOB      : {lob}"
)
print(
    f"Support Model Key : "
    f"{support_model_key}"
)
print(
    f"Spend Model Key   : "
    f"{spend_model_key}"
)
print(
    f"\nQC File:\n"
    f"{qc_output}"
)
print(
    f"\nSelected Variables : "
    f"{len(selected_variables):,}"
)
print(
    f"Post-ADI-Merge Rows : "
    f"{post_merge_qc_rows:,}"
)
print(
    f"Variable-Level QC Rows : "
    f"{variable_qc_df['TOTAL_ROWS'].sum():,}"
)
print(
    f"Flat File Row Count Check : "
    f"{'PASS' if post_merge_qc_rows == len(flat_file) else 'FAIL'}"
)
print(
    f"Total Spend        : "
    f"{total_qc_spend:,.2f}"
)
print(
    f"Total Support      : "
    f"{total_qc_support:,.2f}"
)
print("\nNo flat file was created by this QC script.")
print("\nDone.")
