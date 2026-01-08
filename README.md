# 📄 Sistema Automático de Cotizaciones con Google Sheets

Este proyecto permite generar **cotizaciones automáticas** usando Google Sheets + Google Docs + Apps Script.

## 🚀 Funcionalidades
- Generación automática de documentos desde Google Sheets
- Uso de plantilla con variables {{ }}
- Botón para generar cotización por fila
- Registro automático de fecha y link
- Compatible con PDF
- Ideal para empresas de servicios

## 🧩 Tecnologías
- Google Sheets
- Google Apps Script
- Google Docs
- Google Drive

## 📊 Estructura de la hoja
Encabezados requeridos:
id_cotizacion
empresa_cliente
RUC_cliente
nombre_cliente
correo_cliente
dni_cliente
telefono_cliente
contacto_cargo
dirección_cliente
descripcion
tipo_servicio
total
duracion
fecha_generacion
generar_contrato
link_cotizacion


## ⚙️ Instalación
1. Abrir Google Sheets
2. Extensiones → Apps Script
3. Pegar el código de `Code.gs`
4. Colocar los IDs de la plantilla y carpeta
5. Ejecutar una vez para autorizar permisos

## 🖱 Uso
- Seleccionar una fila
- Presionar el botón **Generar cotización**
- El documento se crea automáticamente

## 🔐 Seguridad
No subir IDs reales ni datos sensibles al repositorio.

## 📈 Próximas mejoras
- Exportar PDF automático
- Generación de contratos
- Envío por correo o WhatsApp
- Integración con CRM / ERP

---

Desarrollado para automatización empresarial 🚀

