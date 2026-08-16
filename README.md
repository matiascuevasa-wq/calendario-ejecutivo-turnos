# Calendario de Turnos — Excelencia Operacional

App de calendario semanal, ejecutivos de turno, sorteo de fines de semana,
estadísticas y accesos por gerencia. Este proyecto está listo para publicarse
en **Firebase Hosting** con el código en **GitHub**.

## 0. Requisitos previos

- Tener [Node.js](https://nodejs.org/) instalado (versión 18 o superior).
- Tener una cuenta de Google (para Firebase).
- Tener una cuenta de [GitHub](https://github.com).
- Instalar Firebase CLI una sola vez en tu computador:
  ```bash
  npm install -g firebase-tools
  ```

## 1. Crear el proyecto de Firebase

1. Ve a [console.firebase.google.com](https://console.firebase.google.com/) → **"Agregar proyecto"**.
2. Ponle un nombre (ej: `calendario-turnos-oems`) y sigue los pasos (puedes
   desactivar Google Analytics, no es necesario).
3. Dentro del proyecto, ve a **"Compilación" → "Firestore Database"** →
   **"Crear base de datos"** → elige **"Modo producción"** → elige una
   región cercana (ej. `us-central` o `southamerica-east1`).
4. Ve a **⚙️ Configuración del proyecto → Tus apps → </> (Web)** y registra
   una app (nombre libre, no necesitas Firebase Hosting marcado aún).
   Copia el objeto `firebaseConfig` que te muestra.
5. Pega esos valores en `src/firebase.js`, reemplazando los `"TU_..."`.

## 2. Probar localmente

```bash
npm install
npm run dev
```

Abre la URL que te muestra (normalmente `http://localhost:5173`) y prueba
que todo funcione: inicia sesión como `Excelencia Operacional` / clave
`Excelencia OEMS`, agrega una gerencia, un ejecutivo, una actividad, etc.
Los datos ya se están guardando en tu Firestore real.

## 3. Subir el código a GitHub

```bash
git init
git add .
git commit -m "Primera versión del calendario de turnos"
```

Crea un repositorio nuevo (vacío, sin README) en GitHub, luego:

```bash
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git branch -M main
git push -u origin main
```

## 4. Conectar y desplegar a Firebase Hosting (manual, la forma más simple)

```bash
firebase login
firebase init hosting
```

Cuando te pregunte:
- **"Use an existing project"** → elige el proyecto que creaste en el paso 1.
- **"What do you want to use as your public directory?"** → escribe `dist`
- **"Configure as a single-page app?"** → `Yes`
- **"Set up automatic builds and deploys with GitHub?"** → puedes decir `No`
  aquí y usar el método manual, o `Yes` para que te configure el paso 5
  automáticamente (te va a pedir autorizar GitHub).
- Si te pregunta si quieres sobreescribir `firebase.json` o `.firebaserc`,
  responde `No` (ya vienen configurados en este proyecto).

Luego, cada vez que quieras publicar cambios:

```bash
npm run build
firebase deploy
```

Al terminar te entrega una URL pública tipo
`https://TU_PROYECTO.web.app` — esa es la que puedes compartir con
cualquier persona de la compañía.

## 5. (Opcional) Despliegue automático desde GitHub

Este proyecto ya incluye `.github/workflows/deploy.yml`, que despliega
automáticamente a Firebase Hosting cada vez que haces `git push` a `main`.
Para activarlo:

1. Reemplaza `TU_PROYECTO_ID` dentro de `.github/workflows/deploy.yml` por
   el ID real de tu proyecto Firebase (lo ves en la URL de la consola o en
   `.firebaserc`).
2. Genera una cuenta de servicio:
   ```bash
   firebase init hosting:github
   ```
   Esto crea automáticamente el secreto `FIREBASE_SERVICE_ACCOUNT` en tu
   repositorio de GitHub por ti.
3. Desde ahí en adelante, cada `git push` a `main` publica solo automáticamente.

## 6. Reglas de seguridad (léelo)

`firestore.rules` deja la base de datos abierta a lectura/escritura para
cualquiera con la URL del proyecto — es el mismo nivel de protección que
tenía la app dentro de Claude.ai (permisos solo del lado del navegador, no
autenticación real de servidor). Es aceptable para un uso interno, pero
si más adelante quieres protección real, lo ideal es migrar el login a
**Firebase Authentication**. Para aplicar las reglas:

```bash
firebase deploy --only firestore:rules
```

## Instalar en el celular (PWA)

Una vez publicada (después del Paso 4), abre la URL `https://tu-proyecto.web.app`
desde el navegador del celular:

- **Android (Chrome):** aparece un aviso "Agregar a pantalla de inicio" solo,
  o desde el menú ⋮ → "Instalar aplicación" / "Agregar a pantalla de inicio".
- **iPhone (Safari):** botón compartir (el cuadrado con flecha hacia arriba)
  → "Agregar a pantalla de inicio".

Queda como un ícono más del celular, se abre a pantalla completa (sin barra
del navegador) y funciona igual que la versión web — los datos son los
mismos, es la misma app, solo un acceso directo más cómodo.

## Estructura del proyecto

```
├── src/
│   ├── App.jsx        ← toda la lógica y las pantallas de la app
│   ├── firebase.js     ← credenciales de tu proyecto Firebase (paso 1)
│   ├── main.jsx
│   └── index.css
├── firebase.json        ← configuración de Hosting + Firestore
├── firestore.rules       ← reglas de acceso a la base de datos
├── .firebaserc            ← ID de tu proyecto Firebase
└── .github/workflows/deploy.yml   ← despliegue automático (opcional)
```

## Usuarios

- **Sesión maestra:** usuario `Excelencia Operacional`, clave `Excelencia OEMS`.
  Controla gerencias, ejecutivos de turno, clasificaciones de actividad,
  sorteo de turnos y accesos de gerentes.
- **Gerentes:** se crean desde Administración → Cuentas de gerentes. Solo
  pueden agregar/editar actividades de su propia gerencia.
- **Invitado:** acceso de solo lectura, sin clave, disponible para
  cualquiera desde el mismo botón de inicio de sesión.
