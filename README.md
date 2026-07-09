
# Data Control Hub - Streamlit MVP V1

Primera interfaz amigable para el MVP de Data Control Hub.

## Qué permite hacer

1. Subir Excel operativo mensual.
2. Subir resultado anterior opcional para conservar gestión.
3. Procesar datos con el motor Python.
4. Ver resumen ejecutivo.
5. Revisar validaciones de datos.
6. Gestionar acciones pendientes en una tabla editable.
7. Descargar Excel resultado.
8. Enviar Excel resultado por correo.

## Estructura

```text
Data_Control_Hub_Streamlit_MVP_V1/
│
├── app.py
├── requirements.txt
├── README.md
│
└── core/
    └── engine.py
```

## Cómo ejecutar localmente

1. Instala Python 3.10 o superior.
2. Abre una terminal en la carpeta del proyecto.
3. Instala dependencias:

```bash
pip install -r requirements.txt
```

4. Ejecuta la app:

```bash
streamlit run app.py
```

5. Se abrirá una ventana en tu navegador.

## Cómo usar la app

1. En el panel lateral, sube el Excel operativo.
2. Opcionalmente sube un resultado anterior.
3. Ajusta el umbral de reparación, por defecto 0.70.
4. Presiona **Procesar archivo**.
5. Revisa:
   - Resumen ejecutivo
   - Validación de datos
   - Acciones pendientes
   - Historial
   - Reglas
6. En `Acciones pendientes`, edita:
   - estado_gestion
   - responsable_manual
   - fecha_compromiso
   - fecha_cierre
   - decision_final
   - comentario_manual
   - evidencia
   - usuario_gestion
7. Presiona **Guardar cambios en sesión**.
8. En `Exportar / Enviar`, descarga o envía el Excel.

## Importante

La interfaz todavía no tiene base de datos.  
Por eso, para conservar gestión mes a mes, debes descargar el Excel resultado y usarlo como "resultado anterior" en la siguiente ejecución.

## Flujo mensual recomendado

```text
Mes 1:
Excel operativo → procesar → resultado → gestionar → descargar

Mes 2:
Excel operativo actualizado + resultado anterior gestionado → procesar → nuevo resultado
```

## Gmail

Para enviar por Gmail no uses tu contraseña normal.  
Debes usar una contraseña de aplicación.
