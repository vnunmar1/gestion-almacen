# 🏭 Sistema de Control de Almacén

Control de entradas y salidas paletizadas con extracción automática de datos desde fotos de albaranes usando **Gemini Vision**, confirmación por **Telegram** y almacenamiento en **Google Sheets**.

## Stack

| Componente | Tecnología |
|---|---|
| Frontend operario | Telegram Bot (móvil/escritorio) |
| Backend / orquestador | Google Apps Script |
| OCR / Extracción IA | Gemini 1.5 Flash (Vision) |
| Base de datos | Google Sheets |
| Deploy | clasp (CLI de Apps Script) |

## Estructura del proyecto

```
gestion-almacen/
├── src/
│   ├── appsscript.json   # Manifiesto del proyecto Apps Script
│   ├── setup.gs          # Creación e inicialización de hojas
│   ├── webhook.gs        # Punto de entrada Telegram (doPost)
│   ├── gemini.gs         # Extracción de datos con Gemini Vision
│   ├── states.gs         # Gestión de estados por usuario
│   └── sheets.gs         # Escritura en Sheets + Telegram API
├── .clasp.json           # Config de clasp (scriptId)
├── .gitignore
└── README.md
```

## Instalación paso a paso

### 1. Prerrequisitos

```bash
npm install -g @google/clasp
clasp login
```

### 2. Crear el Google Sheet

1. Crear un libro nuevo en [Google Sheets](https://sheets.google.com)
2. Copiar el **ID del Sheet** desde la URL:
   `https://docs.google.com/spreadsheets/d/**SHEET_ID**/edit`

### 3. Crear el proyecto Apps Script vinculado al Sheet

```bash
# Desde la carpeta del proyecto
clasp create --type sheets --title "Gestion Almacen" --parentId "TU_SHEET_ID"
```

Esto genera el `.clasp.json` con el `scriptId`. **No subas este archivo al repo** (está en `.gitignore`).

### 4. Subir el código

```bash
clasp push
```

### 5. Ejecutar el setup inicial

1. Abrir el Google Sheet
2. Menú **🏭 Almacén → ⚙️ Inicializar sistema (setup)**
3. Esto crea las 4 hojas: `CONFIG`, `STOCK`, `ENTRADAS`, `SALIDAS`

### 6. Rellenar CONFIG

En la hoja `CONFIG`, rellenar los valores marcados en amarillo:

| Clave | Dónde obtenerlo |
|---|---|
| `TELEGRAM_BOT_TOKEN` | [@BotFather](https://t.me/BotFather) → `/newbot` |
| `GEMINI_API_KEY` | [Google AI Studio](https://aistudio.google.com/app/apikey) |
| `ALLOWED_USERS` | Tu ID de Telegram (usa [@userinfobot](https://t.me/userinfobot)) |

### 7. Publicar como Web App

1. En Apps Script: **Implementar → Nueva implementación**
2. Tipo: `Aplicación web`
3. Ejecutar como: `Yo`
4. Acceso: `Cualquier usuario`
5. Copiar la **URL del Web App** generada
6. Pegarla en `CONFIG → WEBHOOK_URL`

También via clasp:
```bash
clasp deploy --description "v1.0"
```

### 8. Registrar el webhook de Telegram

Menú **🏭 Almacén → 🔄 Registrar webhook de Telegram**

O manualmente en el navegador:
```
https://api.telegram.org/bot{TOKEN}/setWebhook?url={WEBHOOK_URL}
```

### 9. Verificar

Enviar `/start` al bot desde Telegram. Debe responder con el mensaje de bienvenida.

## Flujo de uso

```
Operario envía foto del albarán
        ↓
Telegram → Apps Script (webhook)
        ↓
Gemini Vision extrae: tipo, referencia, cliente, palets, fecha, matrícula
        ↓
Bot muestra resumen con botones: ✅ Confirmar | ✏️ Corregir | ❌ Cancelar
        ↓
Operario confirma → Datos escritos en ENTRADAS o SALIDAS
        ↓
Bot responde con stock actual de esa referencia
```

## Comandos disponibles en Telegram

| Comando | Descripción |
|---|---|
| `/start` | Mensaje de bienvenida y ayuda |
| `/stock` | Ver stock actual de todo el almacén |
| `/ayuda` | Instrucciones de uso |
| Foto | Procesar albarán o carta de porte |

## Hojas de Google Sheets

### STOCK
Stock actual calculado automáticamente con fórmulas `SUMPRODUCT`.

### ENTRADAS / SALIDAS
Historial completo de movimientos con: timestamp, fecha albarán, referencia, cliente, palets, matrícula, observaciones, operario.

### CONFIG
Configuración del sistema. Los campos obligatorios están marcados en amarillo.

## Actualizar código

```bash
clasp push
# Si cambias el deploy:
clasp deploy --deploymentId "TU_DEPLOYMENT_ID" --description "v1.1"
```

## Límites del plan gratuito

| Servicio | Límite gratuito |
|---|---|
| Google Apps Script | 6 min/ejecución, 90 h/día |
| Gemini 1.5 Flash | 1.500 req/día, 15 RPM |
| Google Sheets | 10M celdas |
| Telegram Bot API | Sin límite oficial |

**Coste total: 0 €/mes** para volúmenes habituales de almacén.

## Licencia

MIT
