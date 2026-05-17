# 🎭 Gestión de Tarjetas — Club Deportivo Mitre

WebApp para registrar pedidos de tarjetas anticipadas en peñas folklóricas y cenas.  
Datos almacenados en Google Sheets via Apps Script. Frontend en GitHub Pages.

---

## 📋 Contenido de este repositorio

```
├── index.html    ← La WebApp (va a GitHub Pages)
└── Code.gs       ← El backend (va a Google Apps Script)
```

---

## 🚀 Paso a paso para instalar

### PASO 1 — Crear la Google Sheet

1. Ir a [sheets.google.com](https://sheets.google.com) y crear una hoja nueva.
2. Darle un nombre descriptivo, por ejemplo: **"CDM - Tarjetas Peñas"**.
3. No hace falta crear columnas manualmente, el script las crea solo.

---

### PASO 2 — Crear el Apps Script

1. Dentro de la Google Sheet, ir al menú **Extensiones → Apps Script**.
2. Se abre el editor. Borrar todo el contenido del archivo `Código.gs`.
3. Copiar y pegar **todo el contenido de `Code.gs`** de este repo.
4. Guardar (Ctrl+S o el ícono del diskette).

---

### PASO 3 — Desplegar como Web App

1. En el editor de Apps Script, hacer clic en **"Implementar" → "Nueva implementación"**.
2. Hacer clic en el engranaje ⚙️ junto a "Tipo" y seleccionar **"Aplicación web"**.
3. Completar:
   - **Descripción:** `CDM Tarjetas v1`
   - **Ejecutar como:** `Yo (tu email)`
   - **Quién tiene acceso:** `Cualquier persona`
4. Hacer clic en **"Implementar"**.
5. Si pide autorizar, aceptar todos los permisos (es tu propia hoja).
6. **Copiar la URL** que aparece. Se ve así:
   ```
   https://script.google.com/macros/s/AKfy.../exec
   ```
   ⚠️ Guardar esa URL, la vas a necesitar en el siguiente paso.

---

### PASO 4 — Publicar en GitHub Pages

1. Crear un repositorio en [github.com](https://github.com) (puede ser privado o público).
2. Subir el archivo `index.html` a la raíz del repositorio.
3. Ir a **Settings → Pages** del repositorio.
4. En "Source", seleccionar **"Deploy from a branch"**, rama `main`, carpeta `/ (root)`.
5. Guardar. GitHub Pages te dará una URL pública como:
   ```
   https://tu-usuario.github.io/nombre-repo/
   ```

---

### PASO 5 — Conectar la WebApp

1. Abrir la URL de GitHub Pages en el navegador.
2. Aparece la pantalla de configuración inicial.
3. Pegar la URL del Web App (del Paso 3) y hacer clic en **"Guardar y continuar"**.
4. Ir a la pestaña **⚙️ Config** y configurar:
   - Nombre del evento (ej: "Peña Folklórica — Junio 2025")
   - Fecha del evento
   - Tipos de tarjeta y sus precios (ej: General $5.000 / Especial $4.000 / Menor $2.500)
5. Guardar configuración. ¡Listo!

---

## 📱 Cómo usar la app

### ✏️ Registrar
- Ingresá nombre, tipo de tarjeta y cantidad.
- Indicá si pagó **completo**, dejó una **seña** (ingresás el monto), o queda **pendiente**.
- Elegí el método de pago: efectivo o transferencia.

### 📋 Listado
- Ves todos los registros con su estado de pago.
- Filtrá por nombre o por estado (pagados / con seña / pendientes).
- Editá el estado de pago de cualquier entrada.
- Exportá a **CSV** para abrir en Excel o imprimir el listado.

### 📊 Resumen
- Total de tarjetas y personas registradas.
- Cuántas están pagas, con seña o pendientes.
- Monto total cobrado vs. total por cobrar.
- Desglose por tipo de tarjeta.

### ⚙️ Config
- Editá el nombre/fecha del evento.
- Agregá o modificá tipos de tarjeta y precios.
- Actualizá la URL del Web App si cambia.

---

## 🔄 Actualizar el Apps Script

Si el script cambia en el futuro:

1. En el editor de Apps Script, copiar el nuevo `Code.gs`.
2. Ir a **"Implementar" → "Gestionar implementaciones"**.
3. Editar la implementación existente y aumentar la versión.
4. Guardar. La URL del Web App **no cambia**.

---

## ❓ Problemas frecuentes

**"Error al cargar configuración"**  
→ Verificar que la URL del Web App esté correcta y que el script esté desplegado correctamente.

**La primera vez tarda en responder**  
→ Normal. Google necesita "calentar" el script. La segunda llamada es más rápida.

**Los datos no aparecen después de guardar**  
→ Hacer clic en el botón ↻ (actualizar) arriba a la derecha.

**Quiero usar el mismo script para otro evento**  
→ Simplemente cambiar la configuración del evento en ⚙️ Config. Los registros anteriores quedan en la hoja pero podés archivarlos o borrarlos desde la Google Sheet directamente.
