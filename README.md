# AppEncriptarPDF

Modificar las variables necesarias en App_dummy.config y renombrar a App.config

### Uso

```powershell
.\Encriptador.exe <-option (value)> .. <-option (value)>
  -hide: Oculta la consola.
  -killall: Termina todos los procesos en ejecución de nombre Encriptador.exe.
  -manual: Para ejecutar a demanda.
  -archivo <nombrearchivo>: Procesa el archivo especifico <nombrearchivo>.
  -intervalo <tiempo ms>: Invervalo entre los reintentos en caso .pdf bloqueado
  -limite <tiempo ms>: Limite total de los reintentos en caso .pdf bloqueado
  -rutaorigen <ruta>: Procesa los archivos en la <ruta>
  -rutadestino <ruta>: Deja los .pdf con clave en la <ruta>

  Orden no relevante
```

## Variable en Base de Datos
```SQL
SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (17) --RutaInputPCT
SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (18) --RutaOutputPCT
SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (19) --RutaRejectedPCT
```