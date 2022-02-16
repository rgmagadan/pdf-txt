# pdf-txt

Scripts en Powershell para convertir documentos a texto plano.

- Convert-all-docs pasa a un solo archivo txt todos los documentos de la carpeta en la que se ejecuta.
- Convert-doc convierte a txt el archivo cuya ruta se le pasa como parámetro.
- Para ejecutar más cómodamente los scripts se recomienda crear sendos alias en el profile de Powershell.
- Si no se ejecuta el script, abre PowerShell como administrador y pega lo siguiente para dar los permisos necesarios:
Set-ExecutionPolicy RemoteSigned
