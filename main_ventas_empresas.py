import pandas as pd
import os

# ruta ejecutable 

ruta_ejecutable = os.getcwd()

nombre_archivo = "datos_ventas.xlsx"
ruta_archivo = os.path.join(ruta_ejecutable,"insumos", nombre_archivo )

if __name__ == "__main__":

    try:
        # carge de la tabla de excel en un dataframe utizando un metodo de pandas
        df_ventas_empresas = pd.read_excel(ruta_archivo)

        # seleccion de columnas a trabajar en el ejercicio
        df_ventas_empresas = df_ventas_empresas[["ID_Venta","Fecha","Producto","Cantidad","Precio_Unitario","Total_Venta","Vendedor"]]
                                                
        # verificar nulos en todo el datframe
        nulos = df_ventas_empresas.isna().sum()
        print("columnas con valores nulos:\n" ,nulos[nulos > 0]) 
    
        # coompletar valor de columna Total_venta 
        df_ventas_empresas["Total_Venta"] = df_ventas_empresas["Total_Venta"].fillna(df_ventas_empresas["Cantidad"] * df_ventas_empresas["Precio_Unitario"])
        df_ventas_empresas.to_excel("hola.xlsx", index = False)
        # convertir en enteros 
        # df_ventas_empresas["Total_Venta"] = pd.to_numeric(df_ventas_empresas["Total_Venta"], errors="coerce")
        df_ventas_empresas["Total_Venta"] = [int(x) if pd.notna(x) else "" for x in df_ventas_empresas["Total_Venta"]]
     
        
        # cambio de la fecha de tipo object a datetime
        df_ventas_empresas['Fecha'] = pd.to_datetime(df_ventas_empresas['Fecha'],errors='coerce')

        # Filtrar las ventas del año 2026
        df_ventas_empresas = df_ventas_empresas[(df_ventas_empresas['Fecha'] >= '2023-01-01') &(df_ventas_empresas['Fecha'] < '2024-01-01')]

        # crear columna de mes venta 
        df_ventas_empresas["Mes"] = df_ventas_empresas["Fecha"].dt.month

        df_ventas_empresas = df_ventas_empresas.copy()
        # total de ventas por mes usando funciones group by por las columnas vendedor y mes y sumando las ventas de las agrupaciones
        df_total_ventas = (df_ventas_empresas.groupby(["Vendedor", "Mes"])["Total_Venta"].sum().reset_index())

        print("Total de ventas de vendedor por mes")
        print(df_total_ventas)

        # creacion de dataframes agrupados 

        # se agrupa por Vendedor y se suma el total de ventas
        df_resumen_ventas_x_vendedor = (df_ventas_empresas.groupby(["Vendedor"])["Total_Venta"].sum().reset_index())

        # se agrupa por Mes y se suma el total de ventas
        df_resumen_ventas_mes = (df_ventas_empresas.groupby(["Mes"])["Total_Venta"].sum().reset_index())

        Nombre_archivo = "resumen_ventas.xlsx"
        ruta_archivo = os.path.join(ruta_ejecutable,Nombre_archivo)

        nombre_hoja_1 = "Resumen_Ventas"
        nombre_hoja_2 = "Ventas_Mensuales"

        # para guardar varios dataframes un hojas de excel se crea un objeto con Excelwriter para poder hacer el to_excel al mismo archivo
        with pd.ExcelWriter(ruta_archivo, engine="openpyxl") as writer:
            df_resumen_ventas_x_vendedor.to_excel(writer, sheet_name=nombre_hoja_1, index=False)

            df_resumen_ventas_mes.to_excel(writer, sheet_name=nombre_hoja_2, index=False)

    except Exception as e:

        print(f"Se presenton un error en el proceso: {e}")

   








   