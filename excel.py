import pandas as pd
from datetime import datetime

# Limpiar columnas
def clean_column_name(col):
    if isinstance(col, (int, float)):
        return str(col)
    return col.replace('\n', ' ').strip()


def process_materials(excel_file_path):
    try:
        df = pd.read_excel(
            excel_file_path,
            engine='openpyxl',
            skiprows=4,
            header=0,
            na_values=['', 'NA', 'N/A', 'null', 'NULL'],
            thousands=',',
            decimal='.'
        )

        df.columns = [str(col).strip().upper().replace('\n', ' ') for col in df.columns]

        column_mapping = {
            'NOMBRE': ['NOMBRE'],
            'CLAVE INTERNA': ['CLAVE INTERNA'],
            'CLAVE PROVEEDOR': ['CLAVE PROVEEDOR'],
            'PROVEEDOR': ['PROVEEDOR'],
            'DESCRIPCION': ['DESCRIPCION'],
            'CALIBRE SUPERFICIAL': ['CALIBRE SUPERFICIAL (MIL)', 'CALIBRE SUPERFICIAL'],
            'TIPO ADHESIVO': ['TIPO ADHESIVO'],
            'RESPALDO': ['RESPALDO (MIL)', 'RESPALDO'],
            'CALIBRE TOTAL': ['CALIBRE TOTAL (MIL)', 'CALIBRE TOTAL'],
            'TIPO RECUBRIMIENTO EN PELICULAS': ['TIPO RECUBRIMIENTO EN PELICULAS'],
            'PRESENTACIONES': ['PRESENTACIONES'],
            'DISPONIBILIDAD': ['DISPONIBILIDAD'],
            'OBS': ['OBS'],
            'APROBACION PARA CONTACTO CON ALIMENTOS': ['APROBACION PARA CONTACTO CON ALIMENTOS'],
            'SUSTENTABLES': ['SUSTENTABLES'],
            'DIGITAL': ['DIGITAL'],
            'COMPROMISO DE TIEMPO DE ENTREGA': ['COMPROMISO DE TIEMPO DE ENTREGA'],
            'COSTO PARA COTIZAR': ['COSTO PARA COTIZAR'],
            'MON': ['MON', 'MON.1', 'MON.2'],
            'PRECIO POR M2': ['PRECIO POR M2']
        }

        required_columns = [
            'NOMBRE', 'CLAVE INTERNA', 'DESCRIPCION',
            'CALIBRE SUPERFICIAL', 'CALIBRE TOTAL', 'RESPALDO'
        ]

        available_columns = {}
        for std_name, variants in column_mapping.items():
            for col in df.columns:
                if col in variants:
                    available_columns[std_name] = col
                    break

        missing_columns = [col for col in required_columns if col not in available_columns]
        if missing_columns:
            print(f"\nERROR: Faltan columnas requeridas: {missing_columns}")
            print("Columnas disponibles:", list(df.columns))
            return None

        df = df.rename(columns={v: k for k, v in available_columns.items()})

        selected_columns = [col for col in column_mapping.keys() if col in available_columns]
        df = df[selected_columns]

        numeric_columns = [
            'CALIBRE SUPERFICIAL',
            'CALIBRE TOTAL',
            'RESPALDO',
            'COSTO PARA COTIZAR',
            'PRECIO POR M2',
        ]

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

                if df[col].isna().mean() > 0.5:
                    df[col] = pd.to_numeric(
                        df[col].astype(str).str.replace('[^\d.]', '', regex=True),
                        errors='coerce'
                    )

                df[col] = df[col].fillna(0)

        df = df.dropna(subset=['CLAVE INTERNA'])
        df = df[df['CLAVE INTERNA'].astype(str).str.strip() != '']

        return df

    except Exception as e:
        print(f"\nERROR al procesar el archivo de materiales: {str(e)}")
        return None


def process_die_cutters(excel_file_path):
    try:
        df = pd.read_excel(
            excel_file_path,
            engine='openpyxl',
            skiprows=2,
            header=0,
            dtype=str,
            na_values=['', 'NA', 'N/A', 'null', 'NULL'],
            keep_default_na=False
        )

        expected_columns = [
            'empty1', 'internal_number', 'theets', 'die_cutter_number', 'ubication',
            'width', 'x', 'length', 'tape', 'side_gap', 'upper_gap',
            'step_repetitions', 'repetitions_to_development', 'figure',
            'thousand_ml', 'thousand_area', 'material', 'comments', 'provider', 'entry_date'
        ]

        if len(df.columns) != len(expected_columns):
            expected_columns = expected_columns[:len(df.columns)]

        df.columns = expected_columns
        df = df.dropna(how='all')

        return df

    except Exception as e:
        print(f"Error al procesar suajes: {str(e)}")
        return None


def format_sql_value(value, field_name, is_material=False):
    if pd.isna(value) or str(value).strip().lower() in ('', 'na', 'n/a', 'null', 'nan'):
        return None

    value_str = str(value).strip()

    if is_material and field_name in ['surface_caliber', 'backing_caliber', 'total_caliber', 'price_m2']:
        try:
            return str(float(value_str.replace(',', '.')))
        except:
            return None
    elif not is_material and field_name in ['theets', 'width', 'length', 'tape', 'side_gap',
                                            'upper_gap', 'step_repetitions', 'repetitions_to_development',
                                            'thousand_ml', 'thousand_area']:
        try:
            return str(float(value_str.replace(',', '.')))
        except:
            return None
    elif field_name == 'entry_date':
        try:
            if '/' in value_str:
                parts = value_str.split('/')
                if len(parts) == 3:
                    day, month, year = map(int, parts)
                    if year < 100:
                        year += 2000
                    return f"'{year}-{month:02d}-{day:02d}'"
            return f"'{value_str}'"
        except:
            return None
    else:
        return f"'{value_str.replace("'", "''")}'"


def generate_combined_sql_script(materials_df, die_cutters_df, sql_file_path):
    script_header = f"""
-- Total de registros en materiales.xlsx: {len(materials_df) if materials_df is not None else 0}
-- Total de registros en suajes.xlsx: {len(die_cutters_df) if die_cutters_df is not None else 0}

BEGIN;
SET FOREIGN_KEY_CHECKS = 0;

"""

    insert_statements = []
    skipped_records = {'materials': 0, 'die_cutters': 0}

    if materials_df is not None:
        insert_statements.append("\n-- INSERCIONES PARA TABLA materials\n")
        insert_statements.append("DELETE FROM materials;\n")
        insert_statements.append("ALTER TABLE materials AUTO_INCREMENT = 1;\n")

        for index, row in materials_df.iterrows():
            internal_code = str(row.get('CLAVE INTERNA', '')).strip()
            if not internal_code or internal_code.lower() in ('nan', ''):
                skipped_records['materials'] += 1
                continue

            fields = ['id']
            values = [str(index + 1)]

            field_mappings = {
                'name': row.get('NOMBRE'),
                'provider_code': row.get('CLAVE PROVEEDOR'),
                'internal_code': row.get('CLAVE INTERNA'),
                'description': row.get('DESCRIPCION'),
                'shallow_gauge': row.get('CALIBRE SUPERFICIAL'),
                'total_gauge': row.get('CALIBRE TOTAL'),
                'backup': row.get('RESPALDO'),
                'sizes': row.get('PRESENTACIONES'),
                'availability': row.get('DISPONIBILIDAD'),
                'obs': row.get('OBS'),
                'approval': row.get('APROBACION PARA CONTACTO CON ALIMENTOS'),
                'sustainable': row.get('SUSTENTABLES'),
                'digital': row.get('DIGITAL'),
                'delivery_time': row.get('COMPROMISO DE TIEMPO DE ENTREGA'),
                'quote_cost': row.get('COSTO PARA COTIZAR'),
                'currency1': row.get('MON'),
                'price_currency_2': row.get('PRECIO POR M2'),
                'currency2': row.get('MON'),
                'provider_id': '1',
                'adhesive_type_id': '1',
                'coating_type_id': '1',
                'created_at': 'CURRENT_TIMESTAMP',
                'updated_at': 'NULL',
                'deleted_at': 'NULL'
            }

            for field, value in field_mappings.items():
                if field in ['created_at', 'updated_at', 'deleted_at']:
                    formatted_value = value
                elif field in ['shallow_gauge', 'total_gauge', 'backup', 'quote_cost', 'price_currency_2']:
                    try:
                        if pd.isna(value):
                            formatted_value = 'NULL'
                        else:
                            num = float(value)
                            formatted_value = str(num)
                    except:
                        formatted_value = 'NULL'
                else:
                    formatted_value = format_sql_value(value, field, is_material=True)

                if formatted_value is not None and formatted_value != 'NULL':
                    fields.append(field)
                    values.append(formatted_value)

            if len(fields) > 1:
                insert_stmt = f"INSERT INTO materials ({', '.join(fields)}) VALUES ({', '.join(values)});\n"
                insert_statements.append(insert_stmt)
            else:
                skipped_records['materials'] += 1

    if die_cutters_df is not None:
        insert_statements.append("\n-- INSERCIONES PARA TABLA die_cutters\n")
        insert_statements.append("DELETE FROM die_cutters;\n")
        insert_statements.append("ALTER TABLE die_cutters AUTO_INCREMENT = 1;\n")

        start_id = 1

        for index, row in die_cutters_df.iterrows():
            die_cutter_number = str(row.get('die_cutter_number', '')).strip()
            if not die_cutter_number or die_cutter_number.lower() in ('nan', ''):
                skipped_records['die_cutters'] += 1
                continue

            fields = ['id']
            values = [str(start_id + index)]

            field_mappings = {
                'internal_number': row.get('internal_number'),
                'theets': row.get('theets'),
                'die_cutter_number': row.get('die_cutter_number'),
                'ubication': row.get('ubication'),
                'width': row.get('width'),
                'length': row.get('length'),
                'tape': row.get('tape'),
                'side_gap': row.get('side_gap'),
                'upper_gap': row.get('upper_gap'),
                'step_repetitions': row.get('step_repetitions'),
                'repetitions_to_development': row.get('repetitions_to_development'),
                'thousand_ml': row.get('thousand_ml'),
                'thousand_area': row.get('thousand_area'),
                'comments': row.get('comments'),
                'entry_date': row.get('entry_date'),
                'material_id': row.get('material'),
                'provider_id': row.get('provider'),
                'created_at': 'CURRENT_TIMESTAMP',
                'updated_at': 'NULL'
            }

            for field, value in field_mappings.items():
                formatted_value = format_sql_value(value, field) if field not in ['created_at', 'updated_at'] else value
                if formatted_value is not None:
                    fields.append(field)
                    values.append(formatted_value)

            if len(fields) > 1:
                insert_stmt = f"INSERT INTO die_cutters ({', '.join(fields)}) VALUES ({', '.join(values)});\n"
                insert_statements.append(insert_stmt)
            else:
                skipped_records['die_cutters'] += 1

    script_footer = """
SET FOREIGN_KEY_CHECKS = 1;
COMMIT;
-- Fin del script de actualización
"""

    try:
        with open(sql_file_path, 'w', encoding='utf-8') as sql_file:
            sql_file.write(script_header)
            sql_file.writelines(insert_statements)
            sql_file.write(script_footer)

        print(f"\nScript SQL combinado generado exitosamente en: {sql_file_path}")
        print(
            f"Total registros materiales: {len(materials_df) if materials_df is not None else 0} (omitidos: {skipped_records['materials']})")
        print(
            f"Total registros suajes: {len(die_cutters_df) if die_cutters_df is not None else 0} (omitidos: {skipped_records['die_cutters']})")

    except Exception as e:
        print(f"Error al escribir el archivo SQL: {str(e)}")


if __name__ == "__main__":
    materiales_excel = r'C:\Users\AOC-DES-249\Downloads\materials.xlsx'
    suajes_excel = r'C:\Users\AOC-DES-249\Downloads\suajes.xlsx'
    output_sql = r'C:\Users\AOC-DES-249\Downloads\update_train_db.sql'

    materials_df = process_materials(materiales_excel)
    die_cutters_df = process_die_cutters(suajes_excel)

    generate_combined_sql_script(materials_df, die_cutters_df, output_sql)