# Monitor Inversión Argentina

El monitor se publica en: https://claude.ai/artifact/NzPfwKDM2CxwBHmmkDVCib

Este repositorio ya no aloja la página. Se mantiene porque acá corre el proceso que baja todos los días los datos oficiales que usa el monitor:

- `pipeline/` y `.github/workflows/datos.yml`: descarga diaria (BCRA, INDEC, FMI, riesgo país).
- `data/`: resultado de esa descarga; lo lee la actualización automática del monitor.
- `insumos/`: Excel del relevamiento de agencias de promoción de inversiones.

No borrar el repositorio: sin `data/` el monitor deja de actualizarse.
