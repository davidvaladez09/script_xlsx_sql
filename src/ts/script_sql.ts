import fs from 'fs';
import path from 'path';
import { parse } from 'csv-parse/sync';
import xlsx from 'xlsx';

// Define interfaces for our data structures
interface MaterialRow {
  [key: string]: any;
  'NOMBRE'?: string;
  'CLAVE INTERNA'?: string;
  'CLAVE PROVEEDOR'?: string;
  'PROVEEDOR'?: string;
  'DESCRIPCION'?: string;
  'CALIBRE SUPERFICIAL'?: number | string;
  'TIPO ADHESIVO'?: string;
  'RESPALDO'?: number | string;
  'CALIBRE TOTAL'?: number | string;
  'TIPO RECUBRIMIENTO EN PELICULAS'?: string;
  'PRESENTACIONES'?: string;
  'DISPONIBILIDAD'?: string;
  'OBS'?: string;
  'APROBACION PARA CONTACTO CON ALIMENTOS'?: string;
  'SUSTENTABLES'?: string;
  'DIGITAL'?: string;
  'COMPROMISO DE TIEMPO DE ENTREGA'?: string;
  'COSTO PARA COTIZAR'?: number | string;
  'MON'?: string;
  'PRECIO POR M2'?: number | string;
}

interface DieCutterRow {
  empty1?: string;
  internal_number?: string;
  theets?: string | number;
  die_cutter_number?: string;
  ubication?: string;
  width?: string | number;
  x?: string;
  length?: string | number;
  tape?: string | number;
  side_gap?: string | number;
  upper_gap?: string | number;
  step_repetitions?: string | number;
  repetitions_to_development?: string | number;
  figure?: string;
  thousand_ml?: string | number;
  thousand_area?: string | number;
  material?: string;
  comments?: string;
  provider?: string;
  entry_date?: string;
}

// Helper functions
function cleanColumnName(col: string | number): string {
  if (typeof col === 'number') {
    return col.toString();
  }
  return col.replace('\n', ' ').trim();
}

function formatSqlValue(value: any, fieldName: string, isMaterial: boolean = false): string | null {
  if (value === null || value === undefined || value.toString().trim().toLowerCase() === 'nan') {
    return null;
  }

  const valueStr = value.toString().trim();
  if (valueStr === '' || ['na', 'n/a', 'null'].includes(valueStr.toLowerCase())) {
    return null;
  }

  if (isMaterial) {
    const materialNumericFields = ['surface_caliber', 'backing_caliber', 'total_caliber', 'price_m2'];
    if (materialNumericFields.includes(fieldName)) {
      try {
        const num = parseFloat(valueStr.replace(',', '.'));
        return isNaN(num) ? null : num.toString();
      } catch {
        return null;
      }
    }
  } else {
    const dieCutterNumericFields = [
      'theets', 'width', 'length', 'tape', 'side_gap',
      'upper_gap', 'step_repetitions', 'repetitions_to_development',
      'thousand_ml', 'thousand_area'
    ];
    if (dieCutterNumericFields.includes(fieldName)) {
      try {
        const num = parseFloat(valueStr.replace(',', '.'));
        return isNaN(num) ? null : num.toString();
      } catch {
        return null;
      }
    }
  }

  if (fieldName === 'entry_date') {
    if (valueStr.includes('/')) {
      const parts = valueStr.split('/');
      if (parts.length === 3) {
        let [day, month, year] = parts.map(Number);
        if (year < 100) year += 2000;
        return `'${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}'`;
      }
    }
    return `'${valueStr}'`;
  }

  return `'${valueStr.replace(/'/g, "''")}'`;
}

// Process Materials Excel
function processMaterials(excelFilePath: string): MaterialRow[] | null {
    try {
      const workbook = xlsx.readFile(excelFilePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      
      // Leer datos y limpiar nombres de columnas
      const data: any[] = xlsx.utils.sheet_to_json(sheet, { range: 4, defval: null });
      
      if (!data || data.length === 0) {
        console.log("No data found in materials Excel file");
        return null;
      }
  
      // Obtener y limpiar nombres de columnas
      const firstRow = data[0];
      const originalColumns = Object.keys(firstRow);
      const cleanColumns = originalColumns.map(col => 
        cleanColumnName(col).toUpperCase().replace('\n', ' ').trim()
      );
  
      // Mapeo de columnas esperadas
      const columnMapping: Record<string, string[]> = {
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
      };
  
      // Encontrar mapeo real de columnas
      const availableColumns: Record<string, string> = {};
      for (const [stdName, variants] of Object.entries(columnMapping)) {
        for (let i = 0; i < cleanColumns.length; i++) {
          if (variants.includes(cleanColumns[i])) {
            availableColumns[stdName] = originalColumns[i];
            break;
          }
        }
      }
  
      // Verificar columnas requeridas
      const requiredColumns = [
        'NOMBRE', 'CLAVE INTERNA', 'DESCRIPCION',
        'CALIBRE SUPERFICIAL', 'CALIBRE TOTAL', 'RESPALDO'
      ];
      
      const missingColumns = requiredColumns.filter(col => !availableColumns[col]);
      if (missingColumns.length > 0) {
        console.log(`ERROR: Faltan columnas requeridas: ${missingColumns.join(', ')}`);
        console.log("Columnas disponibles:", cleanColumns);
        return null;
      }
  
      // Procesar cada fila
      const processedData: MaterialRow[] = data.map((row: any) => {
        const newRow: MaterialRow = {};
        
        // Mapear columnas originales a nombres estandarizados
        for (const [stdName, origName] of Object.entries(availableColumns)) {
          if (row[origName] !== undefined && row[origName] !== null) {
            newRow[stdName as keyof MaterialRow] = row[origName];
          }
        }
  
        // Procesar columnas numéricas
        const numericColumns = [
          'CALIBRE SUPERFICIAL',
          'CALIBRE TOTAL',
          'RESPALDO',
          'COSTO PARA COTIZAR',
          'PRECIO POR M2',
        ];
  
        numericColumns.forEach(col => {
          if (newRow[col as keyof MaterialRow] !== undefined) {
            const value = newRow[col as keyof MaterialRow];
            if (typeof value === 'number') {
              // Ya es número, no hacer nada
            } else if (typeof value === 'string') {
              // Intentar convertir string a número
              const numValue = parseFloat(value.replace(',', '').replace(/[^\d.-]/g, ''));
              newRow[col as keyof MaterialRow] = isNaN(numValue) ? 0 : numValue;
            } else {
              newRow[col as keyof MaterialRow] = 0;
            }
          }
        });
  
        return newRow;
      });
  
      // Filtrar filas sin CLAVE INTERNA
      return processedData.filter(row => {
        const internalCode = row['CLAVE INTERNA']?.toString().trim();
        return internalCode && internalCode !== '';
      });
  
    } catch (error) {
      console.log(`ERROR al procesar materiales: ${error}`);
      return null;
    }
  }
  
  // Modificamos la generación SQL para materiales
  function generateMaterialInserts(materialsData: MaterialRow[]): string[] {
    const inserts: string[] = [];
    
    materialsData.forEach((row, index) => {
      const fields: string[] = ['id'];
      const values: string[] = [(index + 1).toString()];
  
      // Mapeo de campos
      const mappings = {
        name: row['NOMBRE'],
        provider_code: row['CLAVE PROVEEDOR'],
        internal_code: row['CLAVE INTERNA'],
        description: row['DESCRIPCION'],
        shallow_gauge: row['CALIBRE SUPERFICIAL'],
        total_gauge: row['CALIBRE TOTAL'],
        backup: row['RESPALDO'],
        sizes: row['PRESENTACIONES'],
        availability: row['DISPONIBILIDAD'],
        obs: row['OBS'],
        approval: row['APROBACION PARA CONTACTO CON ALIMENTOS'],
        sustainable: row['SUSTENTABLES'],
        digital: row['DIGITAL'],
        delivery_time: row['COMPROMISO DE TIEMPO DE ENTREGA'],
        quote_cost: row['COSTO PARA COTIZAR'],
        currency1: row['MON'],
        price_currency_2: row['PRECIO POR M2'],
        currency2: row['MON']
      };
  
      // Procesar cada campo
      for (const [field, value] of Object.entries(mappings)) {
        if (value === undefined || value === null) continue;
  
        let formattedValue: string;
        
        // Campos numéricos especiales
        if (['shallow_gauge', 'total_gauge', 'backup', 'quote_cost', 'price_currency_2'].includes(field)) {
          const numValue = typeof value === 'number' ? value : parseFloat(value.toString().replace(',', ''));
          formattedValue = isNaN(numValue) ? 'NULL' : numValue.toString();
        } else {
          // Campos de texto
          formattedValue = `'${value.toString().replace(/'/g, "''")}'`;
        }
  
        if (formattedValue !== 'NULL') {
          fields.push(field);
          values.push(formattedValue);
        }
      }
  
      // Campos fijos
      fields.push('provider_id', 'adhesive_type_id', 'coating_type_id', 'created_at');
      values.push('1', '1', '1', 'CURRENT_TIMESTAMP');
  
      // Crear sentencia INSERT
      inserts.push(`INSERT INTO materials (${fields.join(', ')}) VALUES (${values.join(', ')});`);
    });
  
    return inserts;
  }

// Process Die Cutters Excel
function processDieCutters(excelFilePath: string): DieCutterRow[] | null {
  try {
    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // Skip first 2 rows and use row 3 as headers
    const data: any[] = xlsx.utils.sheet_to_json(sheet, { range: 2, defval: null });

    if (!data || data.length === 0) {
      console.log("No data found in die cutters Excel file");
      return null;
    }

    const expectedColumns = [
      'empty1', 'internal_number', 'theets', 'die_cutter_number', 'ubication',
      'width', 'x', 'length', 'tape', 'side_gap', 'upper_gap',
      'step_repetitions', 'repetitions_to_development', 'figure',
      'thousand_ml', 'thousand_area', 'material', 'comments', 'provider', 'entry_date'
    ];

    // Rename columns to expected names
    const renamedData: DieCutterRow[] = data.map((row: any) => {
      const newRow: DieCutterRow = {};
      const originalColumns = Object.keys(row);
      
      for (let i = 0; i < Math.min(originalColumns.length, expectedColumns.length); i++) {
        newRow[expectedColumns[i] as keyof DieCutterRow] = row[originalColumns[i]];
      }
      
      return newRow;
    });

    // Filter out completely empty rows
    return renamedData.filter((row: DieCutterRow) => {
      return Object.values(row).some(value => value !== null && value !== '');
    });

  } catch (error) {
    console.log(`Error al procesar suajes: ${error}`);
    return null;
  }
}

// Generate Combined SQL Script
function generateCombinedSqlScript(materialsData: MaterialRow[] | null, dieCuttersData: DieCutterRow[] | null, sqlFilePath: string): void {
  const scriptHeader = `
-- Total de registros en materiales.xlsx: ${materialsData?.length || 0}
-- Total de registros en suajes.xlsx: ${dieCuttersData?.length || 0}

BEGIN;
SET FOREIGN_KEY_CHECKS = 0;

`;

  const insertStatements: string[] = [];
  const skippedRecords = { materials: 0, die_cutters: 0 };

  // Process materials data
  if (materialsData) {
    insertStatements.push("\n-- INSERCIONES PARA TABLA materials\n");
    insertStatements.push("DELETE FROM materials;\n");
    insertStatements.push("ALTER TABLE materials AUTO_INCREMENT = 1;\n");

    materialsData.forEach((row, index) => {
      const internalCode = row['CLAVE INTERNA']?.toString().trim();
      if (!internalCode) {
        skippedRecords.materials++;
        return;
      }

      const fields: string[] = ['id'];
      const values: string[] = [(index + 1).toString()];

      const fieldMappings: Record<string, any> = {
        'name': row['NOMBRE'],
        'provider_code': row['CLAVE PROVEEDOR'],
        'internal_code': row['CLAVE INTERNA'],
        'description': row['DESCRIPCION'],
        'shallow_gauge': row['CALIBRE SUPERFICIAL'],
        'total_gauge': row['CALIBRE TOTAL'],
        'backup': row['RESPALDO'],
        'sizes': row['PRESENTACIONES'],
        'availability': row['DISPONIBILIDAD'],
        'obs': row['OBS'],
        'approval': row['APROBACION PARA CONTACTO CON ALIMENTOS'],
        'sustainable': row['SUSTENTABLES'],
        'digital': row['DIGITAL'],
        'delivery_time': row['COMPROMISO DE TIEMPO DE ENTREGA'],
        'quote_cost': row['COSTO PARA COTIZAR'],
        'currency1': row['MON'],
        'price_currency_2': row['PRECIO POR M2'],
        'currency2': row['MON'],
        'provider_id': '1',
        'adhesive_type_id': '1',
        'coating_type_id': '1',
        'created_at': 'CURRENT_TIMESTAMP',
        'updated_at': 'NULL',
        'deleted_at': 'NULL'
      };

      for (const [field, value] of Object.entries(fieldMappings)) {
        if (field === 'created_at' || field === 'updated_at' || field === 'deleted_at') {
          // Skip adding these to fields array, they're handled specially
          continue;
        }

        let formattedValue: string | null = null;
        
        if (field === 'shallow_gauge' || field === 'total_gauge' || field === 'backup' || 
            field === 'quote_cost' || field === 'price_currency_2') {
          try {
            if (value === null || value === undefined) {
              formattedValue = 'NULL';
            } else {
              const num = typeof value === 'number' ? value : parseFloat(value.toString().replace(',', '.'));
              formattedValue = isNaN(num) ? 'NULL' : num.toString();
            }
          } catch {
            formattedValue = 'NULL';
          }
        } else {
          formattedValue = formatSqlValue(value, field, true);
        }

        if (formattedValue !== null && formattedValue !== 'NULL') {
          fields.push(field);
          values.push(formattedValue);
        }
      }

      // Add the special fields
      fields.push('created_at');
      values.push('CURRENT_TIMESTAMP');
      fields.push('updated_at');
      values.push('NULL');
      fields.push('deleted_at');
      values.push('NULL');

      if (fields.length > 1) {
        insertStatements.push(`INSERT INTO materials (${fields.join(', ')}) VALUES (${values.join(', ')});\n`);
      } else {
        skippedRecords.materials++;
      }
    });
  }

  // Process die cutters data
  if (dieCuttersData) {
    insertStatements.push("\n-- INSERCIONES PARA TABLA die_cutters\n");
    insertStatements.push("DELETE FROM die_cutters;\n");
    insertStatements.push("ALTER TABLE die_cutters AUTO_INCREMENT = 1;\n");

    dieCuttersData.forEach((row, index) => {
      const dieCutterNumber = row['die_cutter_number']?.toString().trim();
      if (!dieCutterNumber) {
        skippedRecords.die_cutters++;
        return;
      }

      const fields: string[] = ['id'];
      const values: string[] = [(index + 1).toString()];

      const fieldMappings: Record<string, any> = {
        'internal_number': row['internal_number'],
        'theets': row['theets'],
        'die_cutter_number': row['die_cutter_number'],
        'ubication': row['ubication'],
        'width': row['width'],
        'length': row['length'],
        'tape': row['tape'],
        'side_gap': row['side_gap'],
        'upper_gap': row['upper_gap'],
        'step_repetitions': row['step_repetitions'],
        'repetitions_to_development': row['repetitions_to_development'],
        'thousand_ml': row['thousand_ml'],
        'thousand_area': row['thousand_area'],
        'comments': row['comments'],
        'entry_date': row['entry_date'],
        'material_id': row['material'],
        'provider_id': row['provider'],
        'created_at': 'CURRENT_TIMESTAMP',
        'updated_at': 'NULL'
      };

      for (const [field, value] of Object.entries(fieldMappings)) {
        let formattedValue: string | null = null;
        
        if (field === 'created_at' || field === 'updated_at') {
          // Skip adding these to fields array, they're handled specially
          continue;
        }

        formattedValue = formatSqlValue(value, field);

        if (formattedValue !== null) {
          fields.push(field);
          values.push(formattedValue);
        }
      }

      // Add the special fields
      fields.push('created_at');
      values.push('CURRENT_TIMESTAMP');
      fields.push('updated_at');
      values.push('NULL');

      if (fields.length > 1) {
        insertStatements.push(`INSERT INTO die_cutters (${fields.join(', ')}) VALUES (${values.join(', ')});\n`);
      } else {
        skippedRecords.die_cutters++;
      }
    });
  }

  const scriptFooter = `
SET FOREIGN_KEY_CHECKS = 1;
COMMIT;
-- Fin del script de actualización
`;

  try {
    fs.writeFileSync(sqlFilePath, scriptHeader + insertStatements.join('') + scriptFooter, 'utf8');

    console.log(`\nScript SQL combinado generado exitosamente en: ${sqlFilePath}`);
    console.log(`Total registros materiales: ${materialsData?.length || 0} (omitidos: ${skippedRecords.materials})`);
    console.log(`Total registros suajes: ${dieCuttersData?.length || 0} (omitidos: ${skippedRecords.die_cutters})`);

  } catch (error) {
    console.log(`Error al escribir el archivo SQL: ${error}`);
  }
}

const BASE_DIR = process.cwd(); 
const EXCEL_FOLDER = path.join(BASE_DIR, 'files');
const OUTPUT_FOLDER = path.join(BASE_DIR, 'sql_output');

if (!fs.existsSync(EXCEL_FOLDER)) {
  fs.mkdirSync(EXCEL_FOLDER, { recursive: true });
}
if (!fs.existsSync(OUTPUT_FOLDER)) {
  fs.mkdirSync(OUTPUT_FOLDER, { recursive: true });
}

const MATERIALS_EXCEL = path.join(EXCEL_FOLDER, 'materiales.xlsx');
const DIE_CUTTERS_EXCEL = path.join(EXCEL_FOLDER, 'suajes.xlsx');
const OUTPUT_SQL = path.join(OUTPUT_FOLDER, 'update_train_db.sql');

function checkFilesExist() {
  const errors: string[] = [];
  
  if (!fs.existsSync(MATERIALS_EXCEL)) {
    errors.push(`No se encontró el archivo: ${MATERIALS_EXCEL}`);
  }
  
  if (!fs.existsSync(DIE_CUTTERS_EXCEL)) {
    errors.push(`No se encontró el archivo: ${DIE_CUTTERS_EXCEL}`);
  }
  
  if (errors.length > 0) {
    console.error('Errores encontrados:');
    errors.forEach(err => console.error(`- ${err}`));
    console.log(`\nPor favor coloca los archivos en: ${EXCEL_FOLDER}`);
    return false;
  }
  
  return true;
}

function main() {
  if (!checkFilesExist()) {
    return;
  }

  console.log(`Procesando archivos desde: ${EXCEL_FOLDER}`);
  console.log(`- Materiales: ${MATERIALS_EXCEL}`);
  console.log(`- Suajes: ${DIE_CUTTERS_EXCEL}`);
  
  const materialsData = processMaterials(MATERIALS_EXCEL);
  const dieCuttersData = processDieCutters(DIE_CUTTERS_EXCEL);

  generateCombinedSqlScript(materialsData, dieCuttersData, OUTPUT_SQL);
  
  console.log(`\nArchivo SQL generado en: ${OUTPUT_SQL}`);
}

main();