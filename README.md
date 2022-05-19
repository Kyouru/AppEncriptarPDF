# AppEncriptarPDF

Visual Studio 2017
dotnet 4.0 > System.Data.OracleClient
Modificar las variables necesarias en App_dummy.config y renombrar a App.config

### Uso

```powershell
.\Encriptador.exe <-option (value)> .. <-option (value)>
  -rutaorigen <ruta>: Procesa los archivos en la <ruta>. Obligatorio.
  -rutadestino <ruta>: Deja los .pdf con clave en la <ruta>. Obligatorio.
  -hide: Oculta la consola.
  -killall: Termina todos los procesos en ejecución de nombre Encriptador.exe.
  -auto: Para que se quede revisando los archivos creados.
  -archivo <nombrearchivo>: Procesa el archivo especifico <nombrearchivo>.
  -intervalo <tiempo ms>: Invervalo entre los reintentos en caso .pdf bloqueado.
  -limite <tiempo ms>: Limite total de los reintentos en caso .pdf bloqueado.

```

## Variable en Base de Datos
```SQL
SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (17) --RutaInputPCT
SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (18) --RutaOutputPCT
SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (19) --RutaRejectedPCT
```