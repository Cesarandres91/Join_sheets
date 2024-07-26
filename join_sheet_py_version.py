import pandas as pd
from datetime import datetime

def parse_date(date_string):
    """Convierte una cadena de fecha en formato dd.mm.yy a un objeto datetime."""
    return datetime.strptime(date_string, '%d.%m.%y')

def consolidate_sheets(input_file, output_file):
    """Consolida datos de múltiples hojas de un archivo Excel en un DataFrame y lo exporta como un nuevo archivo Excel."""
    
    # Leer el archivo de Excel y obtener nombres de las hojas
    xls = pd.ExcelFile(input_file)
    
    # Filtrar las hojas que comienzan con "R_"
    sheet_names = [sheet for sheet in xls.sheet_names if sheet.startswith("R_")]
    
    # Ordenar hojas por fecha (más reciente a más antigua)
    sheet_names.sort(key=lambda name: parse_date(name.replace("R_", "")), reverse=True)
    
    # Diccionario para almacenar los datos consolidados
    data = {}
    
    # Leer y consolidar datos de cada hoja
    for sheet_name in sheet_names:
        # Leer la hoja actual
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Obtener solo las primeras 9 columnas
        df = df.iloc[:, :9]
        
        # Procesar cada fila y agregarla al diccionario data
        for _, row in df.iterrows():
            key = str(row[0]).strip()  # Clave basada en la primera columna sin espacios
            if key not in data:
                data[key] = row.values.tolist()
    
    # Crear DataFrame consolidado a partir del diccionario data
    consolidated_df = pd.DataFrame.from_dict(data, orient='index', columns=df.columns.tolist())
    
    # Añadir columnas de cada hoja ordenada con los datos de la última columna
    for sheet_name in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Crear columna adicional basada en la última columna de la hoja actual
        additional_column = []
        for key in consolidated_df.index:
            matching_row = df[df.iloc[:, 0].astype(str).str.strip() == key]
            if not matching_row.empty:
                additional_column.append(matching_row.iloc[0, -1])
            else:
                additional_column.append("")
        
        # Añadir la columna adicional al DataFrame consolidado
        consolidated_df[sheet_name.replace("R_", "")] = additional_column
    
    # Eliminar duplicados basados en la primera columna
    consolidated_df = consolidated_df[~consolidated_df.index.duplicated(keep='first')]
    
    # Exportar el DataFrame consolidado a un archivo Excel
    consolidated_df.to_excel(output_file, index=False)
    
    print(f"Archivo consolidado guardado como {output_file}")

# Uso del script
input_file = 'input_file.xls'  # Reemplaza con la ruta de tu archivo .xls
output_file = 'Resultado.xls'
consolidate_sheets(input_file, output_file)