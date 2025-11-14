
<div align="center">
   
# ⚙️ Utilidad de Mejora de Instalaciones Semi-automaticas

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![Windows](https://img.shields.io/badge/Platform-Windows%2010%2B-success)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

</div>

## 🌟 Resumen del Proyecto

Esta utilidad ha sido desarrollada para **mejorar y optimizar los procesos de instalaciones semi-automáticas** en entornos empresariales o técnicos. La versión 2.0 introduce nuevas herramientas clave para la gestión de paquetes y la integración con plataformas modernas de despliegue como **Intune**.

El objetivo es simplificar tareas repetitivas y ofrecer un control más robusto y auditable sobre los cambios del sistema y los despliegues de software.

---

## ✨ Características Principales (v2.0)

La versión 2.0 incorpora mejoras significativas enfocadas en el despliegue y la gestión de paquetes:

* **Nueva Pestaña 'Pckgr / Intune':** Funcionalidad dedicada a la gestión de paquetes, ideal para entornos que utilicen **Pckgr** o planifiquen despliegues a través de **Microsoft Intune**.
* **Backup Automático:** Las operaciones realizadas en la pestaña 'Pckgr / Intune' incluyen un robusto sistema de backup automático en múltiples formatos:
    * `Excel`
    * `JSON`
    * `ZIP`
* **Integración Directa con Pckgr:** Permite la integración y la preparación directa de paquetes para despliegues masivos.
* **Barra de Progreso Gráfica:** En la sección **'Actualizaciones'**, se ha añadido una barra de progreso visual para ofrecer una retroalimentación clara e inmediata sobre el estado de la tarea.


## 🛠️ Tecnologías Utilizadas

| Tecnología | Propósito |
| :--- | :--- |
| **Python** | Lenguaje de programación principal. |
| **Tkinter/ttkbootstrap** | Creación de la Interfaz Gráfica de Usuario (GUI). |
| **`os` & `glob`** | Manejo del sistema de archivos, directorios y obtención de metadatos (fechas de modificación). |
| **`datetime`** | Formateo y gestión de las fechas de modificación. |

## 💡 Información General y Propósito

| Detalle | Descripción |
| :--- | :--- |
| **Creador** | LSCF |
| **Propósito** | Mejorar y optimizar las instalaciones semi-automáticas. |
| **Origen** | Idea original de LSCF, con soporte en el desarrollo por Inteligencia Artificial (IA). |

### ⚠️ Aviso de Uso y Responsabilidad

Este *software* está desarrollado para **uso libre**. Sin embargo, la responsabilidad sobre el uso del *software* (incluyendo claves, *ports* y archivos portables) recae **exclusivamente en el usuario final, no en el creador**. Asegúrese de cumplir con todas las licencias y políticas aplicables en su entorno.

---

## 📜 Historial de Cambios (Changelog)

### Versión 2.0

* Nueva pestaña 'Pckgr / Intune' añadida.
* Implementación de backup automático de operaciones (Excel, JSON, ZIP).
* Se agregó una barra de progreso visual en la pestaña 'Actualizaciones'.
* Integración y soporte directo con la herramienta Pckgr para despliegues en Intune.

## 🚀 Instalación y Uso

### Prerrequisitos
- Python 3.8 o superior
- Windows 10/11
- Permisos de administrador (recomendado)

---

## 1. Instalación de Dependencias

1.  **Clona el repositorio**
    ```bash
    git clone [https://github.com/LSCF84/CSM.git](https://github.com/LSCF84/CSM.git)
    cd CSM
    ```
2.  **Instala dependencias**
    ```bash
    pip install -r requirements.txt
    ```
    ### 2. Ejecución

Dado que solo utiliza librerías ya estan isntaladas.

1.  Descarga o clona el archivo `csm.py` en tu máquina.
2.  Ejecuta el *script* desde tu terminal:

    ```bash
    python csm.py
    ```

---

## 👨‍💻 Autor

**LSCF**

## ⚙️ Instalación y Dependencias

Para ejecutar este proyecto, necesitas Python 3.x

## 🤝 ¿Quieres contribuir?

¡Claro! Abre un Issue o un Pull Request. Usa la plantilla al crear un Issue.

---

⭐️ Si te sirvió, ¡dale una estrella al repositorio!
