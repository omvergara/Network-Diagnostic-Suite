# 📡 Suite de Diagnóstico de Red (Powershell)

Herramienta portátil "Todo en Uno" para SysAdmins y Soporte Técnico.  
Monitorea latencia, caídas de servicio y calidad Wi-Fi en tiempo real sin instalación.

![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)
![Powershell](https://img.shields.io/badge/Language-PowerShell-5391FE?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)

## 🚀 Características Principales

*   **Monitor en Vivo:** Gráficas de latencia tipo "Electrocardiograma".
*   **Doble Verificación:** Algoritmo inteligente que evita falsas alarmas por micro-cortes de Wi-Fi.
*   **Portátil:** Es un solo archivo `.ps1`. No requiere instalación ni permisos de administrador local (para funciones básicas).
*   **Configurable:** Carga tus propias IPs de servidores o impresoras mediante un archivo JSON o interfaz gráfica.
*   **Logs Automáticos:** Genera reportes `.csv` en el escritorio con el historial de fallos.

## 🛠️ Instalación y Uso

1.  Descarga el archivo `NetworkMonitor.ps1`.
2.  Haz clic derecho sobre el archivo > **Ejecutar con PowerShell**.
3.  *(Opcional)* Si se cierra de inmediato, abre PowerShell y ejecuta:
    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

## ⚙️ Personalización

Al ejecutar la herramienta por primera vez, se generará un archivo `config_monitor.json` en la misma carpeta. Puedes editarlo para:
*   Cambiar los servidores predeterminados.
*   Ajustar el umbral de lentitud (ms).
*   Configurar alertas por correo electrónico (SMTP).

## 📸 Capturas de Pantalla

*(Aquí subirás tus imágenes después)*

## 📄 Licencia

Este proyecto está bajo la Licencia MIT - eres libre de usarlo, modificarlo y distribuirlo.