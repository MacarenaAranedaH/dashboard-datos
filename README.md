# Dashboard de Análisis de Datos (Streamlit)

Dashboard para analizar hojas de Excel/CSV con varias tablas y formatos: crea
gráficos, calcula totales y tasas, aplica filtros dinámicos y guarda los
datasets más usados.

---

## 📁 Estructura del repositorio

Sube estos archivos a tu repositorio de GitHub con esta estructura:

```
tu-repositorio/
├── Tabla.py               # Aplicación principal
├── requirements.txt       # Dependencias
├── .gitignore             # Archivos que NO se suben
└── .streamlit/
    └── config.toml        # Tema y configuración (opcional)
```

> Importante: `config.toml` debe ir **dentro** de una carpeta llamada `.streamlit`.

---

## 🚀 Desplegar en Streamlit Community Cloud (gratis)

1. Crea una cuenta en [GitHub](https://github.com) y sube este proyecto a un
   repositorio (público o privado).
2. Entra a [share.streamlit.io](https://share.streamlit.io) y crea tu cuenta
   conectándola con GitHub.
3. Pulsa **"New app" -> "Use existing repo"** y completa:
   - **Repository**: tu repositorio.
   - **Branch**: normalmente `main`.
   - **Main file path**: `Tabla.py`
   - **App URL** (opcional): elige un subdominio memorable.
4. Pulsa **"Deploy"**. En unos minutos tendras una URL publica tipo
   `https://tu-subdominio.streamlit.app`.

Cada vez que hagas `git push` con cambios, la app se actualiza automaticamente.

### Privacidad
La app hereda los permisos del repositorio: repo publico -> app publica; repo
privado -> app privada. Puedes cambiarlo desde los ajustes de la app e invitar
personas por correo.

---

## Notas importantes para la version en la nube

- **Los "datasets guardados" NO persisten en la nube.** La app los guarda en la
  carpeta local `saved_datasets/`, que en Community Cloud se borra cada vez que
  la app se reinicia o actualiza. Para uso donde subes el archivo cada vez,
  funciona sin problema; si necesitas que persistan, hay que conectar un
  almacenamiento externo (base de datos, Google Sheets, bucket S3, etc.).
- **Memoria limitada (~1 GB).** Archivos muy grandes (cientos de MB) pueden
  agotar la memoria de la app en el plan gratuito. Conviene filtrar o muestrear.

---

## Ejecutar en local (para probar antes de desplegar)

```bash
pip install -r requirements.txt
streamlit run Tabla.py
```

Se abre en `http://localhost:8501`.
