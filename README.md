# 📊 Control Horario PRO

## 🚀 Descripción

Sistema automatizado de control de asistencia desarrollado en Google Sheets mediante Google Apps Script.

Permite procesar registros de entradas y salidas de empleados, calcular horas trabajadas, detectar diferencias con la jornada laboral y clasificar automáticamente cada día (presente, ausente, licencia, feriado o fin de semana).

Está orientado a optimizar la gestión administrativa, reducir errores manuales y mejorar el control del tiempo laboral.

---

## ⚙️ Funcionalidades

### 🧾 Procesamiento de registros

* Lectura automática de datos desde la hoja activa
* Detección dinámica de columnas (Empleado, Fecha, Hora)
* Agrupación por empleado y por día

### ⏱️ Cálculo de horas trabajadas

* Ordena registros de entrada y salida
* Empareja fichadas automáticamente
* Calcula minutos trabajados por jornada
* Convierte resultados a formato horario (HH:MM)

### 🧠 Limpieza de datos

* Eliminación de registros duplicados o muy cercanos (±1 minuto)
* Prevención de errores comunes en fichadas

### 📅 Generación de calendario completo

* Construye automáticamente el rango de fechas analizado
* Incluye días sin registros para detectar ausencias

### 🇦🇷 Feriados automáticos

* Integración con calendario oficial de Argentina
* Identificación automática de días feriados

### 🏖️ Gestión de licencias

* Lectura desde hoja "LICENCIAS"
* Asignación de estados personalizados:

  * Vacaciones
  * Enfermedad
  * Otros
* Prioridad sobre cualquier otro estado

### 👥 Configuración por empleado

* Lectura desde hoja "HORARIOS"
* Definición de jornada laboral por empleado
* Cálculo automático de diferencias

### 📊 Análisis de diferencias

* Detecta:

  * Horas extra (positivo)
  * Horas faltantes (negativo)
* Suma total de diferencias por empleado

### 🎨 Visualización automática

* Generación de hoja "Resultado"
* Formato visual:

  * 🔵 Separación por empleado
  * 🟢 Verde: horas extra
  * 🔴 Rojo: horas faltantes o ausencias
  * 🟠 Feriados destacados

---

## 🧩 Estructura del sistema

Funciones principales del script:

* `controlHorarioPRO()` → ejecución general del sistema
* `obtenerLicencias()` → lectura de ausencias justificadas
* `obtenerHorasPorEmpleado()` → configuración de jornada laboral
* `limpiarDuplicados()` → limpieza de registros
* `obtenerFeriadosRapido()` → detección de feriados
* `pintar()` → aplicación de formato visual

---

## 🔄 Flujo de funcionamiento

1. Lectura de datos de asistencia
2. Agrupación por empleado y fecha
3. Limpieza de registros duplicados
4. Cálculo de horas trabajadas
5. Comparación con jornada laboral
6. Clasificación del día:

   * Presente
   * Ausente
   * Licencia
   * Feriado
   * Fin de semana
7. Generación de hoja "Resultado"
8. Aplicación de formato automático

---

## 🗂️ Estructura de hojas requerida

### Hoja principal (registros)

| Empleado | Fecha | Hora |
| -------- | ----- | ---- |

---

### Hoja: HORARIOS

| Empleado                               | Horas |
| -------------------------------------- | ----- |
| Define horas laborales por día (ej: 8) |       |

---

### Hoja: LICENCIAS

| Empleado                   | Fecha | Tipo |
| -------------------------- | ----- | ---- |
| Ej: VACACIONES, ENFERMEDAD |       |      |

---

## ▶️ Cómo usar

1. Crear una hoja en Google Sheets
2. Abrir Apps Script
3. Copiar el código del proyecto
4. Crear las hojas necesarias:

   * Registros
   * HORARIOS
   * LICENCIAS
5. Ejecutar la función:

   ```javascript
   controlHorarioPRO()
   ```
6. Revisar la hoja "Resultado"

---

## 💡 Casos de uso

* Control de asistencia de empleados
* Gestión administrativa
* Auditoría de horarios
* Automatización de RRHH

---

## 🛠️ Tecnologías utilizadas

* Google Sheets
* Google Apps Script (JavaScript)

---

## 🚀 Valor del proyecto

Este sistema permite:

* Automatizar el control horario
* Reducir errores manuales
* Mejorar la productividad administrativa
* Escalar el control a múltiples empleados

--- 

## 👨‍💼 Autor

**Pablo Vicego**
Lic. en Administración de Empresas
Especialista en Excel Avanzado y Automatización de procesos

---

<img width="1100" height="596" alt="Captura de pantalla (50)" src="https://github.com/user-attachments/assets/e039761f-77e9-4633-b185-0a0d59ad5e38" />
![Uploading Captura de pantalla (50).png…]()
<img width="982" height="641" alt="Captura de pantalla (52)" src="https://github.com/user-attachments/assets/a0f3d372-0a8d-428e-97fe-7a91c89051c6" />


