# pintruder

Este script para python2.7 se conecta a una hoja de Gooogle SpreadhSheet e introduce
las lineas en el ERP PrimaveraBSS

###Requisitos
Este script para __python2.7__ requiere la instalación de los modulos __gspread__ y __google-api-python-client__ para instalarlos simplemente ejecute:
```
pip install gspread
pip install google-api-python-client
```

###Configuración
Este script se configura desde el fichero __pintruder.ini__

##Origen de datos
Puedes obtener una muestra del documento Google Docs mandándome un email a info@ernesto.es

###Compilación
Puedes compilar este script con py2exe ejecutando:
```
python setup.py py2exe
```

###Expiración
Por defecto oauth2 define una expiracón de una hora, para ampliarlo edite el archivo __pintruder.credentials.json__ y edite la variable:
```
"expires_in": 3600
```

### Licencia
Este script está licenciado como __GPLv3__
