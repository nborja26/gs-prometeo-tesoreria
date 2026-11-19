# Tesorería · Grupo Scout Prometeo
1. Sube tu Excel a `datos/` (hojas: **Movimientos**, **Codigos**, **Presupuesto**; mira `datos/plantilla.xlsx`).
2. GitHub Actions ejecuta `scripts/build_data.py` y actualiza `docs/data.json`.
3. La web (`docs/index.html`) lee ese JSON.

Publicar: Settings → Pages → Deploy from branch → `main` / carpeta `/docs`.
Aviso: GitHub Pages es público aunque el repo sea privado. No subas nombres reales (usa los códigos) o protege la web.
