# Dashboard KPI de tareo

Dashboard estático (HTML + JS) que lee el Excel de tareo en el navegador. No usa servidor ni base de datos: el archivo que cargas no se sube a ningún lado.

## Archivos
- `index.html`: estructura de la página
- `styles.css`: estilos (tema claro/oscuro automático)
- `app.js`: lectura del Excel, filtros, KPIs y comparativas
- `datos-ejemplo.js`: datos de muestra; bórralo (y su `<script>` en `index.html`) si no quieres publicar nombres de operadores

## Columnas que reconoce
`odtcod`/`OP`, `ODTDESCRIP`/`DESCRIPCION`, `FECHA`, `DesOperador`/`OPERADOR`, `MaqDes`/`MAQUINA`, `Proceso`, `minutos`, `PliegosParc`/`CANTIDAD BUENA`, `pliegosparcmal`/`CANTIDAD MALA`

## Publicar en GitHub Pages
1. Sube estos archivos a la raíz de un repositorio.
2. Settings → Pages → Deploy from a branch → `main` / `(root)`.
3. Abre la URL que te da GitHub.

Para probar en local, abre `index.html` con doble clic.
