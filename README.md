# AutoConformador 🧾

**AutoConformador** es una herramienta desarrollada para automatizar la generación de notas formales en entornos administrativos y académicos.

---

## 💡 Descripción

**AutoConformador** es un script en Python que automatiza la generación de documentos `.docx` a partir de datos estructurados en un archivo `.csv`. Su objetivo es facilitar la redacción de notas administrativas formales, ahorrando tiempo en tareas repetitivas dentro de entornos institucionales.

Este script es completamente **editable y adaptable a las necesidades del usuario**. En este ejemplo, está diseñado para generar una **nota que otorga conformidad a una factura o tasa**, vinculada a un suministro, pero puede modificarse fácilmente para otros fines administrativos o contextos similares.

---

## 🚀 Funcionalidades implementadas

- Lectura flexible de archivos `.csv` con múltiples variantes de nombres de columna.
- Limpieza automática de montos y detección del tipo de documento (factura/tasa).
- Generación de uno o varios archivos `.docx` con formato formal.
- Estilo aplicado con fuente **Times New Roman 12pt**.
- Soporte para plantilla de Word (`modelo.docx`) con reemplazo dinámico (v1.2).

---

## 🗂️ Estado del proyecto

> 🌱**Versión 1.0 - Documento único:**  
> Primer prototipo funcional. Genera un único archivo con todas las notas juntas.

> 🧩 **Versión 1.1 - Documentos por fila:**  
> Mejora la estructura, genera múltiples archivos, aplica estilo formal.

> ✔ **Versión 1.2 - Plantilla Word:**  
> Utiliza un archivo `.docx` como base y reemplaza texto dinámicamente conservando el formato.

---

## 📌 Requisitos

- Archivo CSV separado por celdas `;`, con columnas como:
  
  - Número de factura/tasa
  - Monto
  - Proveedor
  - Suministro
- Documento base `modelo.docx` (solo requerido para la v1.2)
  

---
