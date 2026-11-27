# Monitor Inversión Argentina

El monitor se publica en: https://claude.ai/artifact/NzPfwKDM2CxwBHmmkDVCib

También hay una copia en GitHub Pages (`index.html`), pensada para quien necesita descargar los informes de IED en Word y las presentaciones del RIGI sin cuenta de Claude. La actualización automática la vuelve a subir después de cada publicación. La página anterior de redirección quedó como `redireccion_anterior.html`.

El repositorio se mantiene además porque acá corre el proceso que baja todos los días los datos oficiales que usa el monitor:

- `pipeline/` y `.github/workflows/datos.yml`: descarga diaria (BCRA, INDEC, FMI, riesgo país).
- `data/`: resultado de esa descarga; lo lee la actualización automática del monitor.
- `insumos/`: Excel del relevamiento de agencias de promoción de inversiones.

No borrar el repositorio: sin `data/` el monitor deja de actualizarse.
