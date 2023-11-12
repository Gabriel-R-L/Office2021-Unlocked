#  Office 2021 LEGAL y ACTIVADO

## Requisitos previos:
- Tener Python
- Abrir CMD como administrador

## Inicio de instalación
Abrir el archivo office2021activation.py de dos posibles formas:
1) Escribir "office2021activation.py"
2) Escribir "py office2021activation.py"
Sin comillas

## Si no se activa la licencia
Copiar el siguiente código en CMD y dar Enter
```
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:kms.msgang.com
cscript ospp.vbs /act
pause
```

### Fuentes utilizadas
- KMS Keys de todo tipo: https://py-kms.readthedocs.io/en/latest/Keys.html

- GVLK para Office: https://learn.microsoft.com/es-es/deployoffice/vlactivation/gvlks

- IDs de productos Office: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run

- Productos Office con su nombre + gvlk: https://github.com/YerongAI/Office-Tool/blob/main/doc/Tech%20Articles/Products.md 

- Configuración para Office Deployment Tool: https://learn.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options 

- Microsoft Office 2021 Activado LEGAL: https://www.youtube.com/watch?v=jE7WhSMrYho 

- Insertar en XML: https://stackoverflow.com/questions/3095434/inserting-newlines-in-xml-file-generated-via-xml-etree-elementtree-in-python + ChatGPT + CodiumAI

- Código para activar licencia plan B: https://github.com/21Z/Microsoft-Office-2021
