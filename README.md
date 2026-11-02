# Directorio de Utilidades de Powershell

## Arranque automático

Las utilidades se cargan solas al abrir **Windows PowerShell 5.1** porque el perfil del usuario hace dot-source de `init.ps1`:

- Archivo de perfil: `C:\Users\<usuario>\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1`
- Línea que lo activa:

  ```powershell
  # Cargar entorno personalizado de PowerShell
  . "C:\u\powershell\init.ps1"
  ```

Esa línea la escribe el propio `init.ps1`: al ejecutarlo crea `$PROFILE` si no existe y se inyecta en él si todavía no está. Para ver la ruta del perfil de la consola actual: `$PROFILE`.

> **PowerShell 7 (`pwsh`)** usa otro perfil (`Documents\PowerShell\Microsoft.PowerShell_profile.ps1`) que hoy no existe, así que ahí no se cargan. Para activarlas, ejecutar una vez desde `pwsh`: `. C:\U\powershell\init.ps1`

### Flujo de carga

1. `init.ps1` → se asegura de estar en `$PROFILE` y llama a `load.ps1`.
2. `load.ps1` → hace dot-source de **todos** los `lib\*.ps1`, registra las funciones detectadas y crea los alias de `$DiccionarioAlias`.
3. `unload.ps1` → elimina de la sesión las funciones y alias cargados.

### Agregar una utilidad nueva

- Crear `lib\<Nombre>.ps1` con el código **dentro de una `function`**. Cualquier código suelto se ejecuta al abrir la consola (y un `exit` la cerraría).
- Opcional: agregar su alias en `$DiccionarioAlias` de `load.ps1`.
- Recargar en la sesión actual: `. C:\U\powershell\load.ps1` (o abrir una consola nueva).

## Mostrar-Arbol

Alias: `ma`, `arbol`

Muestra el árbol de directorios con íconos por tipo de archivo

## Convert-DocToDocx

Alias: `doc2docx`

Convierte todos los `.doc` de la carpeta actual a `.docx`. Usa `wordconv.exe` de Office si está instalado y, si no, Word por COM (cierra las instancias abiertas de Word).

- `-Overwrite` sobrescribe los `.docx` existentes.
- `-DeleteSource` borra los `.doc` convertidos correctamente.

```powershell
cd C:\ruta\con\docs
doc2docx -Overwrite
```
