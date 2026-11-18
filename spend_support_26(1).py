import pandas as pd

from pathlib import Path

from datetime import datetime, timezone


# ==================================================================

# SETTINGS


# ==================================================================

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

    "LOB_METRIC",

    "MMM_REFRESH_QUARTER"
]

EXCLUDED_DESCRIPTIONS = {

    "ACTUAL",

    "SEASONALITY",

    "PREDICTED",

    "INTERCEPT"

}

TRANSFORMATION_WEEKS = 26


# ==================================================================

# 1. GET FILE PATHS


# ==================================================================

output_path_input = input(

    "\nEnter the full output folder path: "

).strip().strip('"')

ads_path_input = input(

    "Enter the full path of the ADS file: "

).strip().strip('"')

mapping_path_input = input(

    "Enter the full path of the Variable Mapping file: "

).strip().strip('"')

output_folder = Path(output_path_input)

ads_path = Path(ads_path_input)

mapping_path = Path(mapping_path_input)

if not output_folder.exists():

    output_folder.mkdir(parents=True, exist_ok=True)

print(f"\nOutput folder selected: {output_folder}")


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


# ==================================================================

# 3. DETERMINE LOB FROM ADS FILE NAME


# ==================================================================

# Example:

#

# USAA_P&C_Master_ADS.csv

#

# USAA | P&C | Master | ADS

#        ^

#        LOB

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

        f"\nLOB could not be determined from:\n{file_name}"

    )


# ==================================================================

# 4. MODEL KEYS


# ==================================================================

support_model_key = f"{lob} - Support"

spend_model_key = f"{lob} - Spend"

brand_support_model_key = "Brand - Support"

brand_spend_model_key = "Brand - Spend"


# ==================================================================

# 5. OUTPUT PATH


# ==================================================================

output_csv = (

    output_folder /

    f"USAA_{lob}_Spend_Support_Flat_File.csv"

)


# ==================================================================

# 6. DISPLAY INFORMATION


# ==================================================================

print("\n" + "=" * 70)

print("PRODUCTION - SPEND/SUPPORT FLAT FILE")

print("=" * 70)

print(f"ADS File          : {ads_path}")

print(f"Variable Mapping  : {mapping_path}")

print(f"Detected LOB      : {lob}")

print(f"Support Model Key : {support_model_key}")

print(f"Spend Model Key   : {spend_model_key}")

print(f"Output File       : {output_csv}")


# ==================================================================

# 7. READ ADS


# ==================================================================

print("\nReading ADS file...")

ads = pd.read_csv(

    ads_path,

    low_memory=False

)

print(f"ADS rows    : {len(ads):,}")

print(f"ADS columns : {len(ads.columns):,}")


# ==================================================================

# 8. ADS COLUMN STRUCTURE


# ==================================================================

if len(ads.columns) < 3:

    raise ValueError(

        "ADS file must contain at least 3 columns."

    )

period_col = ads.columns[0]

geo_col = ads.columns[1]

ads_variable_columns = list(ads.columns[2:])

print("\nADS structure:")

print(f"Period column       : {period_col}")

print(f"Geography column    : {geo_col}")

print(f"First ADS variable  : {ads_variable_columns[0]}")

print(f"Last ADS variable   : {ads_variable_columns[-1]}")

print(f"Total ADS variables : {len(ads_variable_columns):,}")


# ==================================================================

# 9. CLEAN PERIOD


# ==================================================================

print("\nCleaning dates...")

ads[period_col] = pd.to_datetime(

    ads[period_col],

    errors="coerce"

)

invalid_dates = ads[period_col].isna().sum()

if invalid_dates > 0:

    print(

        f"WARNING: {invalid_dates:,} invalid dates removed."

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

# If 00_00_Total_US is the only geography, retain it.

# Otherwise remove it.

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

if len(non_total_us_geographies) > 0:

    initial_rows = len(ads)

    ads = ads[

        ads[geo_col] != "00_00_Total_US"

    ].copy()

    removed_total_us = (

        initial_rows - len(ads)

    )

    print(

        f"Removed Total US rows : {removed_total_us:,}"

    )

else:

    print("Only 00_00_Total_US geography found.")

    print("Total US will be retained.")


# ==================================================================

# 12. CLEAN GEOGRAPHY VALUE


# ==================================================================

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

ads[geo_col] = ads[geo_col].apply(

    clean_geography

)


# ==================================================================

# 13. READ VARIABLE MAPPING


# ==================================================================

print("\nReading Variable Mapping...")

variable_df = pd.read_excel(

    mapping_path,

    sheet_name="Variable",

    keep_default_na=False

)

model_key_df = pd.read_excel(

    mapping_path,

    sheet_name="Model Key",

    keep_default_na=False

)

print(

    f"Variable mapping rows : {len(variable_df):,}"

)

print(

    f"Model Key rows        : {len(model_key_df):,}"

)


# ==================================================================

# 14. CLEAN MAPPING COLUMN NAMES


# ==================================================================

variable_df.columns = [

    str(col).strip()

    for col in variable_df.columns

]

model_key_df.columns = [

    str(col).strip()

    for col in model_key_df.columns

]

if len(variable_df.columns) < 4:

    raise ValueError(

        "The 'Variable' sheet must contain at least 4 columns."

    )

if len(model_key_df.columns) < 5:

    raise ValueError(

        "The 'Model Key' sheet must contain at least 5 columns."

    )


# ==================================================================

# 15. VARIABLE MAPPING


# ==================================================================

# Variable sheet:

#

# A = Support 3.3

# B = New Spend 3.3

# C = ADI Level 3

# D = Customer Segment

mapping_support_col = variable_df.columns[0]

mapping_spend_col = variable_df.columns[1]

mapping_adi_col = variable_df.columns[2]

mapping_segment_col = variable_df.columns[3]

for col in [

    mapping_support_col,

    mapping_spend_col,

    mapping_adi_col,

    mapping_segment_col

]:

    variable_df[col] = (

        variable_df[col]

        .fillna("")

        .astype(str)

        .str.strip()

    )


# ==================================================================

# 16. CREATE SUPPORT / SPEND LOOKUPS


# ==================================================================

support_lookup = {}

spend_lookup = {}

for _, row in variable_df.iterrows():

    support_name = row[mapping_support_col]

    spend_name = row[mapping_spend_col]

    adi_level_3 = row[mapping_adi_col]

    customer_segment = row[mapping_segment_col]

    if not customer_segment:

        customer_segment = "n/a"

    if support_name:

        support_lookup[support_name] = {

            "adi_level_3": adi_level_3,

            "customer_segment": customer_segment

        }

    if spend_name:

        spend_lookup[spend_name] = {

            "adi_level_3": adi_level_3,

            "customer_segment": customer_segment

        }


# ==================================================================

# 17. MODEL KEY MAPPING


# ==================================================================

# Model Key sheet:

#

# A = Model Key

# C = Model KPI

# D = Customer Segment

# E = Interaction Channel

model_key_mapping = model_key_df.iloc[

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


# ==================================================================

# 18. MODEL KEY LOOKUP


# ==================================================================

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


# ==================================================================

# 19. FIND ALL-ZERO VARIABLES


# ==================================================================

print("\nChecking ADS variables...")

all_zero_variables = []

total_variables = len(

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

        or counter == total_variables

    ):

        print(

            f"Checked "

            f"{counter:,}/"

            f"{total_variables:,}"

        )

all_zero_set = set(

    all_zero_variables

)


# ==================================================================

# 20. SKIP TRANSFORMATION PERIOD


# ==================================================================

unique_weeks = (

    ads[period_col]

    .drop_duplicates()

    .sort_values()

    .tolist()

)

if len(unique_weeks) <= TRANSFORMATION_WEEKS:

    raise ValueError(

        f"ADS contains only "

        f"{len(unique_weeks):,} unique weeks. "

        f"At least "

        f"{TRANSFORMATION_WEEKS + 1:,} weeks are required."

    )

transformation_end_date = (

    unique_weeks[

        TRANSFORMATION_WEEKS - 1

    ]

)

ads = ads[

    ads[period_col] > transformation_end_date

].copy()

remaining_weeks = (

    ads[period_col]

    .drop_duplicates()

    .nunique()

)

print(

    f"Transformation weeks skipped : "

    f"{TRANSFORMATION_WEEKS:,}"

)

print(

    f"Transformation period end    : "

    f"{transformation_end_date.strftime('%m/%d/%Y')}"

)

print(

    f"Weeks remaining for output   : "

    f"{remaining_weeks:,}"

)

print(

    f"Output period start          : "

    f"{ads[period_col].min().strftime('%m/%d/%Y')}"

)

print(

    f"Output period end            : "

    f"{ads[period_col].max().strftime('%m/%d/%Y')}"

)


# ==================================================================

# CALCULATE MMM_REFRESH_QUARTER
# ==================================================================

# Use the latest 13 unique weekly periods remaining after the
# transformation-period filter. Select the calendar quarter with
# the most weeks. If quarters are tied, select the most recent one.

latest_13_weeks = (

    ads[period_col]

    .dropna()

    .drop_duplicates()

    .sort_values()

    .tail(13)

)

if latest_13_weeks.empty:

    raise ValueError(
        "No weekly periods are available to calculate "
        "MMM_REFRESH_QUARTER."
    )

quarter_counts = (

    pd.Series(
        latest_13_weeks.dt.to_period("Q").tolist()
    )

    .value_counts()

)

maximum_quarter_count = quarter_counts.max()

selected_quarter = max(

    quarter

    for quarter, count in quarter_counts.items()

    if count == maximum_quarter_count

)

mmm_refresh_quarter = (

    f"{selected_quarter.year}_Q{selected_quarter.quarter}"

)

print(
    f"MMM refresh quarter            : {mmm_refresh_quarter}"
)


# ==================================================================

# 21. SELECT VARIABLES


# ==================================================================

selected_variables = []

for variable in ads_variable_columns:

    is_brand = "_BRD_" in variable

    # --------------------------------------------------------------

    # SUPPORT

    # --------------------------------------------------------------

    if variable in support_lookup:

        info = support_lookup[variable]

        adi_level_3 = info["adi_level_3"]

        customer_segment = info["customer_segment"]

        if (

            adi_level_3.strip().upper()

            in EXCLUDED_DESCRIPTIONS

        ):

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

        info = spend_lookup[variable]

        adi_level_3 = info["adi_level_3"]

        customer_segment = info["customer_segment"]

        if (

            adi_level_3.strip().upper()

            in EXCLUDED_DESCRIPTIONS

        ):

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


# ==================================================================

# 22. GENERATE FLAT FILE


# ==================================================================

print("\n" + "=" * 70)

print("GENERATING FLAT FILE")

print("=" * 70)

output_chunks = []

total_selected = len(

    selected_variables

)

current_utc_timestamp = (

    datetime.now(timezone.utc)

    .strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]

    + "Z"

)


# ==================================================================

# 23. PROCESS SELECTED VARIABLES


# ==================================================================

for counter, item in enumerate(

    selected_variables,

    start=1

):

    variable = item["ads_variable"]

    model_key = item["model_key"]

    adi_level_3 = item["adi_level_3"]

    variable_type = item["type"]

    is_brand = item["is_brand"]

    # --------------------------------------------------------------

    # Model Key information

    # --------------------------------------------------------------

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

        item["customer_segment"]

    )

    interaction_channel = model_info.get(

        "INTERACTION_CHANNEL",

        "n/a"

    )

    if not mapped_customer_segment:

        mapped_customer_segment = (

            item["customer_segment"]

        )

    # --------------------------------------------------------------

    # Extract required ADS columns

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

    # Convert values

    # --------------------------------------------------------------

    temp["VALUE"] = pd.to_numeric(

        temp["VALUE"],

        errors="coerce"

    ).fillna(0)

    # --------------------------------------------------------------

    # Create output fields

    # --------------------------------------------------------------

    temp["MODEL_KEY"] = model_key

    temp["MODEL_KPI"] = model_kpi

    temp["CUSTOMER_SEGMENT"] = (

        mapped_customer_segment

        if mapped_customer_segment

        else "n/a"

    )

    temp["INTERACTION_CHANNEL"] = (

        interaction_channel

        if interaction_channel

        else "n/a"

    )

    temp["VARIABLE"] = adi_level_3

    # --------------------------------------------------------------

    # Spend / Support

    # --------------------------------------------------------------

    if variable_type == "Spend":

        temp["SPEND"] = temp["VALUE"]

        temp["SUPPORT"] = 0

    else:

        temp["SPEND"] = 0

        temp["SUPPORT"] = temp["VALUE"]

    # --------------------------------------------------------------

    # Volume

    # --------------------------------------------------------------

    temp["VOLUME"] = 0

    # --------------------------------------------------------------

    # WEEK_END

    # --------------------------------------------------------------

    temp["WEEK_END"] = (

        temp["WEEK_END"]

        .dt.strftime("%Y-%m-%d")

    )

    # --------------------------------------------------------------

    # Timestamp

    # --------------------------------------------------------------

    temp["TIMESTAMP_UTC"] = (

        current_utc_timestamp

    )

    # --------------------------------------------------------------

    # LOB_METRIC

    # --------------------------------------------------------------

    temp["LOB_METRIC"] = (

        "Brand"

        if is_brand

        else f"{lob} Spend-Support"

    )

    temp["MMM_REFRESH_QUARTER"] = mmm_refresh_quarter

    # --------------------------------------------------------------

    # Select final columns

    # --------------------------------------------------------------

    temp = temp[

        OUTPUT_COLUMNS

    ].copy()

    output_chunks.append(temp)

    # --------------------------------------------------------------

    # Progress

    # --------------------------------------------------------------

    if (

        counter % 10 == 0

        or counter == total_selected

    ):

        percent = (

            counter /

            total_selected *

            100

        )

        print(

            f"Progress: "

            f"{counter:,}/"

            f"{total_selected:,} "

            f"({percent:.1f}%)"

        )


# ==================================================================

# 24. COMBINE OUTPUT


# ==================================================================

if output_chunks:

    final_output = pd.concat(

        output_chunks,

        ignore_index=True

    )

else:

    final_output = pd.DataFrame(

        columns=OUTPUT_COLUMNS

    )


# ==================================================================

# 25. POST-ADI LEVEL 3 MERGE


# ==================================================================

# Multiple ADS variables mapping to the same ADI Level 3

# are combined within the same output grain.

aggregation_columns = [

    "MODEL_KEY",

    "MODEL_KPI",

    "GEOGRAPHY",

    "CUSTOMER_SEGMENT",

    "INTERACTION_CHANNEL",

    "VARIABLE",

    "WEEK_END",

    "TIMESTAMP_UTC",

    "LOB_METRIC",

    "MMM_REFRESH_QUARTER"
]

if not final_output.empty:

    final_output = (

        final_output

        .groupby(

            aggregation_columns,

            sort=False,

            dropna=False,

            as_index=False

        )[

            [

                "SPEND",

                "SUPPORT",

                "VOLUME"

            ]

        ]

        .sum()

    )


# ==================================================================

# 26. FINAL COLUMN ORDER


# ==================================================================

final_output = final_output[

    OUTPUT_COLUMNS

]


# ==================================================================

# 27. SAVE CSV


# ==================================================================

print("\nSaving production flat file...")

final_output.to_csv(

    output_csv,

    index=False

)


# ==================================================================

# 28. FINAL MESSAGE


# ==================================================================

print("\n" + "=" * 70)

print("PRODUCTION COMPLETED")

print("=" * 70)

print(

    f"LOB                 : {lob}"

)


print(

    f"Support Model Key   : {support_model_key}"

)

print(

    f"Spend Model Key     : {spend_model_key}"

)

print(

    f"Variables Generated : "

    f"{len(selected_variables):,}"

)

print(

    f"Output Rows         : "

    f"{len(final_output):,}"

)

print(

    f"\nOutput File:\n"

    f"{output_csv}"

)

print("\nDone.")
