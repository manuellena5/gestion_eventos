# 🎭 Gestión de Tarjetas — Club Deportivo Mitre

WebApp para registrar pedidos de tarjetas anticipadas en peñas folklóricas y cenas.  
Datos almacenados en Google Sheets via Apps Script. Frontend en GitHub Pages.

---

## 📋 Contenido de este repositorio

```
├── index.html    ← La WebApp (va a GitHub Pages)
├── config.js     ← URL del backend y versión de la app
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

1. Abrir el archivo `config.js` y pegar la URL del Web App (del Paso 3) en `API_URL`:
   ```js
   const API_URL = 'https://script.google.com/macros/s/AKfy.../exec';
   ```
2. Subir `config.js` junto con `index.html` al repositorio.
3. Abrir la URL de GitHub Pages en el navegador: la app ya entra directo, sin pedir la URL.
   > Al actualizar el script usar **Implementar → Administrar implementaciones → Editar (lápiz) → Versión: Nueva versión → Implementar**. Así la URL no cambia y no hay que tocar `config.js`.
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

1. En el editor de Apps Script, reemplazar todo el contenido por el nuevo `Code.gs` y guardar.
2. Ir a **Implementar → Administrar implementaciones**.
3. Tocar el lápiz (**Editar**) de la implementación que ya existe.
4. En **Versión** elegir **Nueva versión** y tocar **Implementar**.
5. La URL del Web App **no cambia** (no hace falta tocar `config.js`).

> ⚠️ No usar "Nueva implementación": eso crea otra URL y la app dejaría de conectarse.

---

## 🧪 Versión 2 — probar el backend nuevo y migrar (en curso)

La versión 2 agrega hojas nuevas (`Usuarios`, `Eventos`, `Productos`, `Pedidos`, `Items`, `Pagos`, `Contactos`), que se crean solas. Las hojas viejas (`Registros`, `Config`, `Movimientos` y las archivadas `Registros X` / `Movimientos X`) **no se borran ni se modifican**, y la app actual sigue funcionando.

**Probar primero en una copia de la planilla:** Archivo → Hacer una copia (copia también el script). En la copia, abrir Extensiones → Apps Script y, eligiendo la función en el desplegable de arriba y tocando **Ejecutar**:

1. `testV2` — corre todas las pruebas. En el registro tiene que decir "Todas las pruebas pasaron".
2. `borrarDatosDePrueba` — borra lo que creó la prueba.
3. `migrarDesdeV1` — pasa los datos viejos a la versión 2. El registro muestra, por cada evento, los totales de antes y de ahora y "¿Coinciden los totales? SÍ". Se puede ejecutar de nuevo: no duplica nada.
4. (Opcional) `prepararPruebaCupo` — necesita que la copia esté implementada como Web App. Muestra un código para pegar en la consola del navegador (F12) que manda dos pedidos a la vez por el último lugar: solo uno tiene que pasar. Después, `borrarDatosDePrueba`.

**Usuarios:** en la hoja `Usuarios` escribir un nombre por fila (columna `nombre`). `pin` vacío = entra sin contraseña; para poner un PIN, escribirlo en `pin`. Para sacar a alguien, poner `FALSE` en `activo`. El `id` se completa solo.

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
