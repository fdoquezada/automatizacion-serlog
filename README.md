# 📦 Proyecto de Automatización - AGUNSA

Sistema web para procesamiento y validación de órdenes de transporte punto a punto, permitiendo comparar archivos Excel de pedidos y generar reportes automáticos.

## 🎯 Descripción del Proyecto

Esta aplicación facilita la gestión y validación de datos de transporte mediante:
- **Carga de archivos Excel**: Procesa archivos de órdenes y pedidos
- **Validación automática**: Cruza datos entre múltiples fuentes
- **Generación de reportes**: Exporta resúmenes en formato Excel
- **Interfaz intuitiva**: Diseño responsivo con Bootstrap 5

## 📋 Características Principales

✅ Carga y procesamiento de archivos Excel (.xlsx)  
✅ Validación de referencias externas y estado de pedidos  
✅ Cálculo automático de toneladas y verificaciones  
✅ Generación de reportes descargables  
✅ Filtrado por fecha de movimiento  
✅ Interfaz responsiva (mobile-friendly)  
✅ Soporte multiidioma (interfaz en español)

## 🗂️ Estructura del Proyecto

```
automatizacionv1/
├── index.html                    # Página de inicio
├── Punto_a_Punto_Proyecto.html   # Aplicación principal
├── gowit.html                    # Página complementaria
├── css/                          # Estilos si están separados
├── js/
│   └── gowit.js                  # Lógica principal de la aplicación
├── img/                          # Imágenes y recursos
└── README.md                     # Este archivo
```

## 🚀 Cómo Usar

### Requisitos Previos
- Navegador web moderno (Chrome, Firefox, Safari, Edge)
- Archivos Excel con el formato especificado

### Pasos para Ejecutar

1. Abre `Punto_a_Punto_Proyecto.html` en tu navegador
2. Carga el archivo principal con las órdenes de transporte
3. Carga el archivo de "Pedidos" para validación
4. (Opcional) Filtra por fecha si deseas un rango específico
5. Presiona el botón "Cargar" para procesar los datos
6. Visualiza los resúmenes y descarga el reporte

### Formato de Archivos Excel Esperado

**Archivo Principal (Punto_a_Punto):**
- Contiene datos de órdenes de transporte
- Estructura a ser procesada por la aplicación

**Archivo de Pedidos:**
- Columna A: Disponibilidad
- Columna B (Índice 1): Referencia Externa (clave de cruce)
- Columna E (Índice 4): Verificación (Toneladas)
- Columna H (Índice 7): Estado Pedido
- Valores de estado esperados: "EN EJECUCION", "SERVICIO FINALIZADO", "ANULADO"

## 💻 Tecnologías Utilizadas

- **HTML5**: Estructura semántica
- **CSS3**: Estilos y diseño responsivo
- **JavaScript (Vanilla)**: Lógica de aplicación
- **Bootstrap 5.3.2**: Framework CSS
- **Bootstrap Icons**: Íconos
- **XLSX.js**: Procesamiento de archivos Excel
- **Google Fonts**: Tipografía (IBM Plex)

## 📊 Funcionalidades Principales

### Procesamiento de Datos
- Lee archivos Excel manteniendo formato original
- Mapea servicios activos por referencia externa
- Valida estados de pedidos con priorización
- Calcula resúmenes automáticos

### Generación de Reportes
- Exporta datos procesados a Excel
- Permite descargar resultados formateados
- Filtrado por rango de fechas

### Interfaz de Usuario
- Carga de archivos mediante drag & drop (si está implementado)
- Visualización de previsualización de datos
- Botones de acción intuitivos
- Información de procesamiento en tiempo real

## 🔍 Código Principal

El archivo `js/gowit.js` contiene:
- Manejo de eventos de carga de archivos
- Lógica de procesamiento de datos Excel
- Mapeo de servicios activos
- Generación de reportes
- Funciones de validación

## 📝 Notas Importantes

- La aplicación requiere que los archivos Excel sigan el formato establecido
- Los estados de pedido se priorizan: "EN EJECUCION" > "SERVICIO FINALIZADO" > "ANULADO"
- Los datos se procesan en el cliente (no requiere servidor)
- Se mantiene la información de encabezado originalizada

## 🤝 Contribuciones

Para reportar bugs o sugerir mejoras, contacta al equipo de desarrollo.

## 📞 Contacto

Proyecto desarrollado para AGUNSA  
Empresa de transporte y logística

---

**Última actualización**: Abril 2026  
**Versión**: 1.0
