# Script PowerShell: Abrir Chrome con la última sesión y auto-refrescar cada minuto

Este repositorio contiene un script de PowerShell que:

- Abre Google Chrome usando tu sesión/perfil habitual (cookies, logins, etc.).
- Navega a una URL específica (puedes cambiarla en un parámetro o directamente en el código).
- Refresca la página automáticamente cada minuto (configurable).

> Nota: Para lograr el refresco automático, el script usa WScript.Shell para activar la ventana de Chrome y enviar la tecla F5. Esto implica que, en el momento del refresco, Chrome puede pasar al frente y tomar el foco unos milisegundos.

## Requisitos

- Windows con PowerShell 5.1 o PowerShell 7+.
- Google Chrome instalado.
- Windows Script Host habilitado (para `WScript.Shell`).

## Archivos

- `Open-Chrome-LastSession-Refresh.ps1`: Script principal.

## Uso rápido

1. Abre PowerShell.
2. (Opcional) Permite ejecución de scripts para tu usuario:
   - `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`
   - O bien ejecutar una sola vez con bypass: `powershell -ExecutionPolicy Bypass -File .\Open-Chrome-LastSession-Refresh.ps1`
3. Ejecuta el script indicando la URL y (opcional) el intervalo:

```
# En el directorio del repo
./Open-Chrome-LastSession-Refresh.ps1 -Url "https://tu-sitio.com" -IntervalSeconds 60
```

Si no especificas `-Url`, el script usa `https://www.example.com`. El intervalo por defecto es 60 segundos.

## Cómo cambiar la URL en el código

Si prefieres dejarla fija en el script, edita el valor por defecto del parámetro `Url` al inicio del archivo `Open-Chrome-LastSession-Refresh.ps1`:

```
param(
    [string]$Url = "https://www.ejemplo.com",
    [int]$IntervalSeconds = 60
)
```

## Notas y limitaciones

- El script abre Chrome en una ventana nueva sobre tu perfil actual (misma sesión). Si Chrome ya está abierto, es posible que el proceso que se arranca termine rápido y la nueva ventana pertenezca al proceso ya existente; el script intentará localizar una ventana principal de Chrome para refrescarla.
- Para refrescar, el script activa la ventana de Chrome y envía F5; si estás trabajando en otra ventana en el momento del refresco, puede robar el foco brevemente.
- Para detener el script, cierra la ventana de Chrome abierta o interrumpe PowerShell con Ctrl+C.

## Solución de problemas

- "No se pudo iniciar Google Chrome": verifica que Chrome esté instalado. Rutas probadas automáticamente:
  - `%ProgramFiles%\Google\Chrome\Application\chrome.exe`
  - `%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe`
  - `%LocalAppData%\Google\Chrome\Application\chrome.exe`
- Si PowerShell bloquea la ejecución de scripts, usa `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` o el modo `-ExecutionPolicy Bypass` para una ejecución puntual.

## Personalización avanzada

- Intervalo de refresco: cambia `-IntervalSeconds` (por defecto, 60).
- Si deseas que siempre use un perfil concreto (por ejemplo `Default` o `Profile 1`), puedes modificar los argumentos de inicio en la función `Start-ChromeWindow` para incluir `--profile-directory=Default`.
