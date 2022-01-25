# pdf-txt

Script en Powershell para convertir documentos a texto plano.

- Coloca el ps1 en cualquier directorio del disco.
- En el bat, especifica la ruta del ps1 (las comillas deben quedar puestas).
- Copia el bat en c:/windows.
- En PowerShell, sitúate en cualquier carpeta que contenga los documentos a convertir.
- Ejecuta: pdf-txt.bat.
- En la carpeta se creará file.txt con el texto de todos los documentos.
- Si no se ejecuta el script, abre PowerShell como administrador y pega lo siguiente para dar los permisos necesarios:
Set-ExecutionPolicy RemoteSigned
