import pandas as pd
import pyodbc  # For database connection
from openpyxl import load_workbook  # For reading Excel files


def fetch_final_table_schema(connection_string, table_name):
    query = f"""
    SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = '{table_name}'
    """
    conn = pyodbc.connect(connection_string)
    schema_df = pd.read_sql_query(query, conn)
    conn.close()
    return schema_df



def extract_loader_schema(excel_path, sheet_name):
    workbook = load_workbook(filename=excel_path, data_only=True)
    sheet = workbook[sheet_name]
    data = sheet.values
    columns = next(data)
    df = pd.DataFrame(data, columns=columns)
    workbook.close()
    return df


field_mapping = {
    "P1_field1": ("fruits", "fruit_column1"),
    "P1_field2": ("fruits", "fruit_column2"),
    "P1_field6": ("vehicles", "vehicle_column1"),
    # Add the rest...
}


def validate_formats(final_schema, loader_schema, field_mapping):
    validation_results = []
    for final_field, (loader_table, loader_field) in field_mapping.items():
        final_format = final_schema.loc[final_schema['COLUMN_NAME'] == final_field]
        loader_format = loader_schema.loc[loader_schema['COLUMN_NAME'] == loader_field]

        if final_format.empty or loader_format.empty:
            validation_results.append((final_field, loader_field, "Field not found"))
            continue

        final_type = final_format.iloc[0]['DATA_TYPE']
        loader_type = loader_format.iloc[0]['DATA_TYPE']

        if final_type != loader_type:
            validation_results.append((final_field, loader_field, f"Type mismatch: {final_type} vs {loader_type}"))
    return validation_results


def generate_report(validation_results, output_path):
    results_df = pd.DataFrame(validation_results, columns=["Final Field", "Loader Field", "Issue"])
    results_df.to_csv(output_path, index=False)


# Configuration
db_connection_string = "Your_Connection_String"
final_table_name = "P1"
excel_path = "path_to_loader_schema.xlsx"
output_path = "validation_report.csv"

# Fetch final table schema
final_schema = fetch_final_table_schema(db_connection_string, final_table_name)

# Fetch loader table schemas
fruits_schema = extract_loader_schema(excel_path, "fruits")
vehicles_schema = extract_loader_schema(excel_path, "vehicles")

# Merge schemas into a dictionary
loader_schemas = {
    "fruits": fruits_schema,
    "vehicles": vehicles_schema,
}

# Validate formats
validation_results = validate_formats(final_schema, loader_schemas, field_mapping)

# Generate report
generate_report(validation_results, output_path)
