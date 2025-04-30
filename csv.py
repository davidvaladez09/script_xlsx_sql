import pandas as pd
from datetime import datetime

#Limpiar columnas
def clean_column_name(col):
    return col.replace('\n', ' ').strip()

def process_materials(csv_file_path):
    try:
        #Leer CSV de materiales
        df = pd.read_csv(
            csv_file_path,
            encoding='utf-8',
            skiprows=4,
            header=0,
            dtype=str,
            na_values=['', 'NA', 'N/A', 'null', 'NULL'],
            thousands=',',
            decimal='.'
        )

        df.columns = [clean_column_name(col) for col in df.columns]

        column_mapping = {
            'CLAVE INTERNA': ['CLAVE INTERNA', 'CLAVE  INTERNA', 'CLAVE_INTERNA'],
            'NOMBRE': ['NOMBRE'],
            'CLAVE PROVEEDOR': ['CLAVE PROVEEDOR', 'CLAVE_PROVEEDOR'],
            'PROVEEDOR': ['PROVEEDOR'],
            'DESCRIPCION': ['DESCRIPCION'],
            'CALIBRE SUPERFICIAL (Mil)': ['CALIBRE SUPERFICIAL (Mil)', 'CALIBRE SUPERFICIAL'],
            'TIPO ADHESIVO': ['TIPO ADHESIVO', 'ADHESIVO'],
            'RESPALDO (Mil)': ['RESPALDO (Mil)', 'RESPALDO'],
            'CALIBRE TOTAL (Mil)': ['CALIBRE TOTAL (Mil)', 'CALIBRE TOTAL'],
            'TIPO RECUBRIMIENTO EN PELICULAS': ['TIPO RECUBRIMIENTO EN PELICULAS', 'RECUBRIMIENTO'],
            'PRESENTACIONES': ['PRESENTACIONES'],
            'DISPONIBILIDAD': ['DISPONIBILIDAD'],
            'OBS': ['OBS', 'OBSERVACIONES'],
            'APROBACION PARA CONTACTO CON ALIMENTOS': ['APROBACION PARA CONTACTO CON ALIMENTOS',
                                                       'APROBACION ALIMENTOS'],
            'SUSTENTABLES': ['SUSTENTABLES'],
            'DIGITAL': ['DIGITAL'],
            'COMPROMISO DE TIEMPO DE ENTREGA': ['COMPROMISO DE TIEMPO DE ENTREGA', 'TIEMPO ENTREGA'],
            'COSTO PARA COTIZAR': ['COSTO PARA COTIZAR', 'COSTO COTIZAR'],
            'MON': ['MON', 'MONEDA'],
            'PRECIO POR M2': ['PRECIO POR M2', 'PRECIO M2']
        }

        required_columns = ['NOMBRE', 'CLAVE INTERNA', 'DESCRIPCION']
        missing_columns = [col for col in required_columns
                           if not any(mapped_col in df.columns for mapped_col in column_mapping.get(col, [col]))]

        if missing_columns:
            print(f"Error: Faltan columnas requeridas: {missing_columns}")
            return None

        #Renombrar columnas
        final_columns = {}
        for standard_name, variants in column_mapping.items():
            for variant in variants:
                if variant in df.columns:
                    final_columns[standard_name] = variant
                    break

        df = df.rename(columns={v: k for k, v in final_columns.items()})
        df = df[list(final_columns.keys())]

        #Limpiar datos
        df = df.dropna(subset=['CLAVE INTERNA'])
        df = df[df['CLAVE INTERNA'].str.strip() != '']

        return df

    except Exception as e:
        print(f"Error al procesar materiales: {str(e)}")
        return None


def process_die_cutters(csv_file_path):
    try:
        #Leer archivo CSV de suajes
        df = pd.read_csv(
            csv_file_path,
            encoding='utf-8',
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

    #Columnas con decimales
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
    #Manejar la fecha
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
    # Manejo de texto
    else:
        return f"'{value_str.replace("'", "''")}'"


def generate_combined_sql_script(materials_df, die_cutters_df, sql_file_path):
    script_header = f"""-- Script combinado de actualización para materiales y suajes
-- Generado automáticamente el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
-- Total de registros en materiales.csv: {len(materials_df) if materials_df is not None else 0}
-- Total de registros en suajes.csv: {len(die_cutters_df) if die_cutters_df is not None else 0}

BEGIN;
SET FOREIGN_KEY_CHECKS = 0;

"""

    insert_statements = []
    skipped_records = {'materials': 0, 'die_cutters': 0}

    if die_cutters_df is not None:
        insert_statements.append("\n-- INSERCIONES PARA TABLA die_cutters\n")
        insert_statements.append("DELETE FROM die_cutters;\n")
        insert_statements.append("ALTER TABLE die_cutters AUTO_INCREMENT = 1;\n")

        for index, row in die_cutters_df.iterrows():
            die_cutter_number = str(row.get('die_cutter_number', '')).strip()
            if not die_cutter_number or die_cutter_number.lower() in ('nan', ''):
                skipped_records['die_cutters'] += 1
                continue

            fields = ['id']
            values = [str(index + 1)]

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
    materiales_csv = r'C:\Users\AOC-DES-249\Downloads\materiales.csv'
    saujes_csv = r'C:\Users\AOC-DES-249\Downloads\suajes.csv'
    output_sql = r'C:\Users\AOC-DES-249\Downloads\update_train_db.sql'

    materials_df = process_materials(materiales_csv)
    die_cutters_df = process_die_cutters(saujes_csv)

    generate_combined_sql_script(materials_df, die_cutters_df, output_sql)