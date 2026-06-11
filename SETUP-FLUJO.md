# Setup — Resumen Sous Chef automático

Pipeline: Office Script filtra y agrega la base pesada dentro de Excel Online → Power Automate lo corre todos los días y guarda `sous-chef-data.json` → el dashboard lee ese JSON (KB en vez de MB). Si el JSON no existe, el dashboard cae automáticamente al xlsx completo (lento pero funciona).

## 1. Crear el Office Script (una sola vez)

1. Abrir **Evolucion ventas Facturado.xlsx** en Excel Online (el navegador, no la app de escritorio).
2. Pestaña **Automatizar** → **Nuevo script**.
3. Borrar el contenido de ejemplo y pegar todo `office-script-resumen-sous-chef.ts`.
4. Renombrarlo a **Resumen Sous Chef** y guardar.
5. Probar con **Ejecutar**: en el panel de resultado tiene que aparecer un string JSON largo. Si tira error de columnas, avisame con el mensaje exacto.

## 2. Crear el flujo de Power Automate (una sola vez)

En [make.powerautomate.com](https://make.powerautomate.com) → **Crear** → **Flujo de nube programado**:

1. **Recurrencia**: cada 1 día, a las 07:00 (zona horaria Buenos Aires).
2. Acción **Excel Online (Business) → Ejecutar script**:
   - Ubicación: el site *Sector Administración y Finanzas*
   - Biblioteca: Documentos
   - Archivo: `Documentos/01- Contabilidad/02- Contabilidad mensual/01- Ventas/02- Reportes venta/Evolucion ventas Facturado.xlsx`
   - Script: **Resumen Sous Chef**
3. Acción **SharePoint → Crear archivo**:
   - Dirección del sitio: *Sector Administración y Finanzas*
   - Ruta de carpeta: `Documentos/01- Contabilidad/02- Contabilidad mensual/01- Ventas/02- Reportes venta`
   - Nombre: `sous-chef-data.json`
   - Contenido: el **Resultado** (`result`) del paso anterior (contenido dinámico)
4. Guardar y **Probar** manualmente. Verificar que el JSON apareció en la carpeta.

Notas:
- Si "Crear archivo" falla la segunda vez porque el archivo ya existe, reemplazarlo por la acción **SharePoint → Enviar una solicitud HTTP a SharePoint** con método `POST` y URI:
  `_api/web/GetFolderByServerRelativeUrl('/sites/SectorAdministracionyFinanzas/Documentos compartidos/Documentos/01- Contabilidad/02- Contabilidad mensual/01- Ventas/02- Reportes venta')/Files/add(url='sous-chef-data.json',overwrite=true)` y el resultado del script en el cuerpo.
- Si no tenés permiso de escritura en esa carpeta de finanzas, guardalo en cualquier carpeta donde sí puedas escribir y actualizá `DATA_JSON_URL` en `index.html` con la nueva ruta (es la ruta relativa a la raíz de la biblioteca del site).

## 3. Dashboard

`index.html` ya intenta primero `sous-chef-data.json` (constante `DATA_JSON_URL`). No requiere más cambios. En la consola del navegador se ve qué fuente usó: `[ventas] usando resumen JSON pre-procesado` o el warning de fallback.

## Pendiente conocido

El archivo de consumos (`consumosEneroalprincipiodeMayo.xlsx`) sigue siendo un snapshot manual congelado en mayo — CCC, retención y churn salen de ahí. El mismo esquema (Office Script sobre el archivo de consumos vigente + paso extra en este flujo) lo resuelve cuando haya un archivo de consumos que se actualice.
