#  Lab Process Automation (VBA/Excel)

[![VBA](https://img.shields.io/badge/VBA-Automation-yellow?logo=microsoft-excel&logoColor=white)]()
[![Lab](https://img.shields.io/badge/Laboratory-Systems-blue)]()

Repositorio de automatización de flujos de trabajo desarrollado para **AVALQUIMICO SAS**. Esta herramienta optimiza la entrada de datos técnicos mediante macros de eventos en Excel.

## Impacto de Negocio
* **Eficiencia Operativa:** Automatiza la generación de sub-parámetros técnicos (Grasa, Humedad, Cenizas), eliminando la carga manual de filas por cada análisis proximal detectado.
* **Integridad de Datos:** Normaliza prefijos de areas (QH, QICL, AAMT, etc.) para asegurar que los los archivos anidados coincidan con la hoja de calculo de cada area.

## 🛠️ Características Técnicas
* **Arquitectura Modular:** El código separa la lógica de normalización de la de inserción de datos para facilitar el mantenimiento.
* **Manejo de Errores:** Implementa `On Error GoTo` para evitar bloqueos del sistema ante datos inesperados.
* **Optimización de Procesos:** Utiliza `Application.ScreenUpdating` para mejorar la velocidad de ejecución en hojas con miles de registros.

## 🚀 Cómo Implementar
1. Abrir el editor de VBA en Excel (ALT + F11).
2. Pegar el contenido de `LabAutomation.vba` en el objeto `Sheet` (Hoja) correspondiente.
3. El sistema se ejecutará automáticamente ante cualquier cambio en las columnas F o G.

## ✒️ Autor
**Giancarlo** - *Ingeniero de Sistemas en formación (8vo semestre) - UNICUCES*
