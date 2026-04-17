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
✅ **Gestión de activación y cierre de viajes** (activacion.html)  
✅ **Vista previa de datos con estadísticas en tiempo real** (index.html)  
✅ **Búsqueda y filtrado avanzado de datos**  
✅ **Exportación de datos procesados a Excel**  
✅ **Navegación integrada entre módulos**  
✅ **Página de error 404 con redirección automática**

## 🔒 Mejoras de Seguridad y Arquitectura

### Separación de Concerns
- ✅ **HTML puro**: Sin estilos inline (`style=`)
- ✅ **CSS externalizado**: Todos los estilos en archivos `.css` separados
- ✅ **JavaScript externalizado**: Sin scripts inline en los HTML
- ✅ **Permite Content Security Policy (CSP)**: Rechaza scripts no autorizados

### Diseño Responsivo
- ✅ **Mobile-first**: Funciona correctamente en smartphones, tablets y desktops
- ✅ **Layout flexible**: Usa flexbox y grid para adaptarse a cualquier pantalla
- ✅ **Footer sticky**: Se posiciona correctamente al pie de la página
- ✅ **Tablas adaptables**: Se ajustan automáticamente en pantallas pequeñas
- ✅ **Navegación responsive**: Menú adaptado a móviles

### Mejor Rendimiento
- ✅ **Caching del navegador**: Los archivos CSS y JS se cachean localmente
- ✅ **Separación de responsabilidades**: Facilita mantenimiento y debugging
- ✅ **Código modular**: Cada página tiene sus propios estilos y lógica

## 🗂️ Estructura del Proyecto

```
automatizacionv1/
├── index.html                    # Página de inicio con preview y estadísticas
├── Punto_a_Punto_Proyecto.html   # Aplicación principal de punto a punto
├── gowit.html                    # Página de reportes Gowit
├── activacion.html               # Gestión de activación y cierre de viajes
├── Historicoalarmas.html         # Filtro de alarmas por fecha
├── error.html                    # Página de error 404
├── css/
│   ├── index.css                 # Estilos para la página principal
│   ├── punto-a-punto-proyecto.css # Estilos para la aplicación principal
│   ├── gowit.css                 # Estilos para reportes Gowit
│   ├── activacion.css            # Estilos para activación
│   ├── historicoalarmas.css      # Estilos para filtro de alarmas
│   └── error.css                 # Estilos para página de error
├── js/
│   ├── index.js                  # Lógica para index.html (preview, stats, búsqueda)
│   ├── punto-a-punto-proyecto.js # Lógica para aplicación principal (filtros, exportación)
│   ├── gowit.js                  # Lógica para reportes Gowit
│   ├── gowit-setup.js            # Configuración de feedback visual para archivos
│   ├── activacion.js             # Lógica para activación y cierre de viajes
│   └── historicoalarmas.js       # Lógica para filtro de alarmas
├── img/                          # Imágenes y recursos (logos, íconos)
└── README.md                     # Este archivo
```

## 🚀 Cómo Usar

### Requisitos Previos
- Navegador web moderno (Chrome, Firefox, Safari, Edge)
- Archivos Excel con el formato especificado

### Navegación del Sistema
- **index.html**: Vista previa y estadísticas generales
- **Punto_a_Punto_Proyecto.html**: Procesamiento principal de punto a punto
- **gowit.html**: Generación de reportes Gowit
- **activacion.html**: Gestión de activación y cierre de viajes
- **Historicoalarmas.html**: Filtro de alarmas por fecha
- **error.html**: Página de error (acceso automático en caso de URL inválida)

### Pasos para Ejecutar

1. Abre `index.html` en tu navegador para una vista general
2. Navega a `Punto_a_Punto_Proyecto.html` para la aplicación principal
3. Carga el archivo principal con las órdenes de transporte
4. Carga el archivo de "Pedidos" para validación
5. (Opcional) Filtra por fecha si deseas un rango específico
6. Presiona el botón "Cargar" para procesar los datos
7. Visualiza los resúmenes y descarga el reporte
8. Usa `activacion.html` para gestionar estados de viajes
9. Usa `Historicoalarmas.html` para filtrar alarmas por fecha

### Usando Historicoalarmas
1. Carga un archivo Excel que contenga una columna "Fecha Alarma"
2. Selecciona la fecha a filtrar usando el selector de fecha
3. Presiona el botón "Filtrar"
4. Los datos se mostrarán en una tabla responsiva
5. La tabla se puede desplazar horizontalmente en dispositivos pequeños

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
- **Navegación entre módulos del sistema**
- **Vista previa con estadísticas en tiempo real** (activos/inactivos)
- **Búsqueda y filtrado avanzado en tablas**
- **Gestión de activación y cierre de viajes**
- **Manejo de errores con redirección automática**

### Gestión de Estados
- **Activación de viajes**: Análisis de tiempos de inicio y fin
- **Cierre de servicios**: Control de estados operativos
- **Validación de estados**: Priorización EN EJECUCION > SERVICIO FINALIZADO > ANULADO

## 🔍 Código Principal

- **js/index.js**: Manejo de preview, estadísticas, búsqueda y filtros en la página principal
- **js/punto-a-punto-proyecto.js**: Lógica de procesamiento de datos Excel, filtros avanzados y exportación
- **js/gowit.js**: Procesamiento de reportes Gowit y validación de datos
- **js/gowit-setup.js**: Configuración de feedback visual para carga de archivos
- **js/activacion.js**: Lógica para gestión de activación y cierre de viajes
- **js/historicoalarmas.js**: Filtrado de alarmas por fecha desde archivos Excel

## 📝 Notas Importantes

- La aplicación requiere que los archivos Excel sigan el formato establecido
- Los estados de pedido se priorizan: "EN EJECUCION" > "SERVICIO FINALIZADO" > "ANULADO"
- Los datos se procesan en el cliente (no requiere servidor)
- Se mantiene la información de encabezado originalizada

## 🤝 Contribuciones

Para reportar bugs o sugerir mejoras, contacta al equipo de desarrollo.

## 📞 Contacto

Proyecto desarrollado para AGUNSA 
por Fernando Quezada
Email-> **fdoquezadapuno@gmail.com** 
Empresa de transporte y logística

---

**Última actualización**: Abril 2026  
**Versión**: 1.1  
**Estado**: Mejoras de seguridad y responsividad implementadas
