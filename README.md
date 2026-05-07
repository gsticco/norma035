# Agente Norma 035 — Guía de despliegue

## Archivos incluidos
```
norma035/
├── server.js          ← Backend (Node.js + Express)
├── package.json       ← Dependencias
└── public/
    └── index.html     ← Frontend (chat)
```

---

## Opción A — Probar en tu computadora (local)

### Requisitos
- Node.js 18 o mayor → https://nodejs.org
- Una API key de Anthropic → https://console.anthropic.com

### Pasos

1. Copia la carpeta `norma035/` a donde quieras en tu computadora.

2. Abre la terminal dentro de esa carpeta:
   ```bash
   cd norma035
   npm install
   ```

3. Crea un archivo llamado `.env` con tu API key:
   ```
   ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxx
   ```

4. Instala dotenv para leer el .env:
   ```bash
   npm install dotenv
   ```

5. Agrega esta línea al INICIO de `server.js`:
   ```js
   require("dotenv").config();
   ```

6. Inicia el servidor:
   ```bash
   npm start
   ```

7. Abre http://localhost:3000 en tu navegador.

---

## Opción B — Desplegar gratis en Railway

Railway es una plataforma gratuita (hasta cierto uso) para alojar apps Node.js.

### Pasos

1. Crea una cuenta en https://railway.app (es gratis)

2. Sube la carpeta a GitHub:
   - Crea un repositorio en https://github.com
   - Sube los archivos (arrastra la carpeta o usa git)

3. En Railway:
   - Haz clic en "New Project" → "Deploy from GitHub"
   - Conecta tu repo
   - Railway detecta automáticamente que es Node.js

4. Agrega la variable de entorno:
   - Ve a tu proyecto → pestaña "Variables"
   - Agrega: `ANTHROPIC_API_KEY` = `sk-ant-xxxxxxxxxxxxx`

5. Railway le asigna una URL pública automáticamente (algo como `tu-app.up.railway.app`).

6. ¡Listo! Comparte esa URL con tus clientes.

---

## Costo estimado

- Railway: gratis hasta ~500 horas/mes de uso
- Anthropic API: ~$0.003 por diagnóstico completo (5 preguntas)
  - Con $5 de crédito inicial puedes hacer ~1,600 diagnósticos gratis

---

## Personalización rápida

Para cambiar las preguntas o el comportamiento del agente, edita la variable
`SYSTEM_PROMPT` en `server.js`. No necesitas tocar el frontend.
