# ✈️ Conquista tu Zona en Aniversario - Escuadrón de Combate (Quest)

Tablero gamificado interactivo estilo portaaviones / escuadrón de combate militar para el seguimiento de cumplimiento de metas retail de las 11 Zonas y 94 PDVs de Colombia.

## 🚀 Despliegue en Vercel

Este proyecto es una SPA (Single Page Application) estática optimizada que no requiere ningún proceso de compilación (`build`).

1. **Subir a GitHub:**
   - Crear un repositorio nuevo en GitHub (por ejemplo `quest-escuadron`).
   - Ejecutar el archivo `SUBIR_A_GITHUB.bat` o usar la terminal git.
2. **Conectar en Vercel:**
   - Ir a [vercel.com/new](https://vercel.com/new).
   - Importar el repositorio `quest-escuadron`.
   - **Framework Preset**: `Other`.
   - **Root Directory**: `./`.
   - Hacer clic en **Deploy**.

## 🛡️ Niveles de Acceso y Seguridad

- **👁️ Modo Espectador (Libre / Sin Contraseña):**
  - Cualquier líder, PDV o colaborador puede ingresar libremente.
  - Visualización 3D interactiva en pista, altímetro, ranking de zonas y tiendas, tabla de clasificación y filtro por zona.
  - **Bloqueado**: No puede cargar archivos Excel ni descargar bases de datos CSV.

- **👑 Modo Maestro (Administrador):**
  - **Usuario:** `Santiago`
  - **Contraseña:** `0814`
  - **Permisos:** Habilita el botón de carga "Nuevo Reporte" (subida inteligente de Excel con soporte para COs, nombres y zonas) y el botón de descarga "Descargar CSV".

## 📦 Estructura del Proyecto

```text
quest-escuadron/
├── index.html           # Aplicación completa (HUD, 3D, Web Audio, SheetJS)
├── quest_logo.png       # Logo oficial de Quest
├── vercel.json          # Reglas de enrutamiento estático para Vercel
├── package.json         # Metadatos del proyecto
├── SUBIR_A_GITHUB.bat   # Script de 1-clic para sincronización con GitHub
└── README.md            # Documentación
```
