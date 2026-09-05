<div align="center">

# 📄 Editor PDF

### Convierte archivos PDF a Word desde una aplicación de escritorio para Windows.

[![Flutter](https://img.shields.io/badge/Flutter-Desktop-02569B?style=for-the-badge&logo=flutter&logoColor=white)](https://flutter.dev/)
[![Dart](https://img.shields.io/badge/Dart-3.12%2B-0175C2?style=for-the-badge&logo=dart&logoColor=white)](https://dart.dev/)
[![Windows](https://img.shields.io/badge/Windows-Desktop-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Word](https://img.shields.io/badge/Microsoft%20Word-Exportación-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white)](https://www.microsoft.com/microsoft-365/word)

`PDF → DOCX` &nbsp; `Flutter` &nbsp; `Windows` &nbsp; `Uso práctico`

</div>

---

## 💡 ¿Qué hace?

**Editor PDF** es una aplicación de escritorio creada con Flutter para convertir archivos PDF en documentos editables de Word (`.docx`).

Permite seleccionar un PDF desde el equipo o descargarlo desde una URL, elegir dónde guardar el resultado y abrir directamente el documento convertido.

## ✨ Funciones principales

| | Función | Descripción |
|:--:|---|---|
| 📁 | **Seleccionar PDF** | Carga un archivo PDF guardado localmente. |
| 🔗 | **Descargar desde URL** | Obtiene un PDF desde un enlace web. |
| 🔄 | **Convertir a Word** | Genera un documento editable en formato `.docx`. |
| 🗂️ | **Elegir ubicación** | Permite definir dónde guardar el documento convertido. |
| 📖 | **Abrir Word** | Abre el resultado directamente en Microsoft Word. |
| 📂 | **Ver carpeta** | Accede rápidamente a la carpeta del documento generado. |
| 🖥️ | **Aplicación de escritorio** | Diseñada para usarse en Windows. |

## 🚀 Cómo usarla

1. Abre la aplicación `Editor PDF`.
2. Selecciona un PDF local o pega la URL de un archivo.
3. Presiona `Descargar` si usas una URL.
4. Presiona `Convertir a Word (.docx)`.
5. Elige dónde guardar el documento.
6. Abre el archivo en Word o revisa su carpeta.

## 🧰 Tecnologías

- **Flutter** para la interfaz de escritorio.
- **Dart** para la lógica de la aplicación.
- **file_picker** para seleccionar archivos y ubicaciones.
- **HTTP** para descargar archivos PDF desde una URL.
- **PowerShell y Microsoft Word** para realizar la conversión a `.docx` en Windows.

## ⚙️ Requisitos

- Windows.
- Flutter SDK 3.12 o superior para ejecutar desde el código fuente.
- Microsoft Word instalado para completar la conversión.

## ▶️ Ejecutar desde el código fuente

```bash
flutter pub get
flutter run -d windows
```

## 📦 Generar la aplicación de Windows

```bash
flutter build windows
```

El resultado se genera dentro de la carpeta `build/windows/`.

## 🧩 Estructura del proyecto

```text
editor_pdf/
|-- lib/                    # Código principal de la aplicación
|-- test/                   # Pruebas
|-- windows/                # Configuración nativa de Windows
|-- pubspec.yaml            # Dependencias y configuración Flutter
|-- pubspec.lock            # Versiones bloqueadas
|-- analysis_options.yaml   # Reglas de análisis Dart
|-- README.md               # Documentación
`-- .gitignore              # Archivos locales excluidos
```

## 🔒 Privacidad

Los archivos PDF se procesan en el equipo. Cuando se utiliza una URL, la aplicación se conecta al sitio de origen para descargar el archivo antes de convertirlo.

## 👩‍💻 Autora

**Isabel Figueroa**

Ingeniera enfocada en automatización, datos, Power BI, Tableau, Python e interfaces web funcionales.

---

<div align="center">

⭐ Una herramienta práctica para llevar documentos PDF a Word.

</div>

