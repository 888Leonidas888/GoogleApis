# Google APIs para VBA

![VBA](https://img.shields.io/badge/VBA-4c9c4c?style=flat&logo=microsoft&logoColor=white)
![API Google Drive](https://img.shields.io/badge/Api%20Google%20Drive-v3-cccccc?style=flat&logo=googledrive&logoColor=white)
![API Google Sheets](https://img.shields.io/badge/Api%20Google%20Sheets-v4-441133?style=flat&logo=googlesheets&logoColor=white)
![OAuth 2.0](https://img.shields.io/badge/OAuth%202.0-Authorization-blue?style=flat&logo=auth0&logoColor=white)


Consuma las APIs de Google Drive y Sheets desde **VBA** con esta DLL escrita en **C#** para **VBA**. 

Eh tratado de abarcar la mayor cantidad de operaciones que le permiten cada API,
desde listar archivos, cargar archivos, descargar archivos, crear carpetas, eliminar archivos, compartir archivos, etc. Y para google sheets, leer y escribir en hojas de cálculo, 
crear hojas de cálculo, eliminar hojas de cálculo, actualizar hojas de cálculo,etc.

Esta DLL también provee una clase llamada `FlowOauth` que se encarga de gestionar el flujo de autenticación y autorización de **OAuth 2.0** para acceder a los recursos de Google Drive y Google Sheets, por lo que no tendrás que preocuparte por la autenticación y autorización de tus aplicaciones.

Los métodos expuestos se basan en la documentación oficial de Google, por lo que si deseas hacer una consulta más específica, te recomiendo que visites los siguientes enlaces:

- [Google Drive API](https://developers.google.com/drive/api/v3/reference)
- [Google Sheets API](https://developers.google.com/sheets/api/reference/rest)
- [Oauth 2.0](https://developers.google.com/identity/protocols/oauth2)


> [!NOTE]
> Dentro de la documentación oficial de Google, encontrarás un apartado donde probar cada método, te recomiendo que pases por ahí antes de echar código.

## Hay mucho por desarrollar

Eh tratado de cubrir la mayor parte de los métodos, sin embargo aún queda mucho por cubrir, si quieres colaborar con el desarrollo de este repositorio haz un fork y sube una PR para discutirlo.

> [!NOTE]
> La función de carga de recursos cubren **carga simple**, **carga multiparte** recomendada para la carga `<= 5mb` revisa para mas detalles [subir datos de archivos](https://developers.google.com/drive/api/guides/manage-uploads?hl=es-419). Cargas mayores no estan soportadas por ahora.

## Tabla de contenido

1. [Instalación](#instalación)
2. [Activar referencias](#activar-referencias)
3. [Configuración de entorno en Google](#configuración-de-entorno-en-google)
4. [Guardar credenciales de acceso](#guardar-credenciales-de-acceso)
5. [Probar FlowOauth y generar el token de acceso](#probar-flowoauth-y-generar-el-token-de-acceso)
6. [Ejemplo de uso](#ejemplo-de-uso)
7. [Recursos adicionales](#Recursos-adicionales)

## Instalación

Descargue, descomprima y ejecute el archivo **setup.exe** para instalar la DLL.

[![Descargar Instalador](https://img.shields.io/badge/⬇-Descargar%20ZIP-green?style=for-the-badge)](https://github.com/888Leonidas888/GoogleApis/releases/download/v1.0.0/setup.zip)


## Activar referencias

Antes de hacer uso, debes asegurarte de tener activadas las siguientes referencias, una vez abierto el archivo desde cual usarás esta DLL te 
saltará una advertencia pidiendo que actives las macros, acepta para continuar, una vez habilitada las macros 
presiona `Alt` + `F11` para ir al VBE. En la barra de menú seleciona **Herramientas** -> **Referencias** -> 
**Examinar**, busca la carpeta donde descargaste y selecciona el siguiente archivo **GoogleApis.tlb**, deberás tener marcada estas referencias:

+ GoogleApis
+ Microsoft Scripting Runtime



> [!NOTE]
> Aparte de las referencias mencionadas líneas arriba, también se debe contar con el siguiente módulo 
[JsonConverter.bas v2.3.1](https://github.com/VBA-tools/VBA-JSON/tree/master), este módulo te facilitará 
la lectura y escritura de archivos **.json**.



## Configuración de entorno en Google

Posiblemente este sea uno de los pasos mas tediosos a seguir pero tomese su tiempo para leerlo detenidamente, pronto agregaré un videotutorial de como hacerlo, pero por ahora siga los pasos en los enlaces o visite [Desarrolla en Google Workspace](https://developers.google.com/workspace/guides/get-started?hl=es_419).

1. [Crea un proyecto de Google Cloud](https://developers.google.com/workspace/guides/create-project?hl=es-419)
2. [Habilita las APIs que deseas usar](https://developers.google.com/workspace/guides/enable-apis?hl=es-419)
3. [Obtén información sobre cómo funcionan la autenticación y autorización](https://developers.google.com/workspace/guides/auth-overview?hl=es-419)
4. [Configura el consentimiento de OAuth](https://developers.google.com/workspace/guides/configure-oauth-consent?hl=es-419)
5. [Crea credenciales de acceso](https://developers.google.com/workspace/guides/create-credentials?hl=es-419)

## Guardar credenciales de acceso

[Las credenciales de acceso](https://developers.google.com/workspace/guides/create-credentials?hl=es-419#api-key) obtenidas debes guardarlas en el directorio **credentials** (no es obligatorio) con extensión **json**, al comienzo solo tendrás 2 archivos; el primero para la [Clave API](https://developers.google.com/workspace/guides/create-credentials?hl=es-419#api-key) y el segundo [ID de cliente de OAuth](https://developers.google.com/workspace/guides/create-credentials?hl=es-419#oauth-client-id)

- **Clave de API:** Guardalo de la siguiente manera, esto es obligatorio, de lo contrario la instancia de `FlowOauth` no podrá encontrar este valor, nombra al archivo como mejor convengas:

```json
{
  "your_api_key": "AIzaSiAsOpGUEW5oS_A6cPkMFLonxGy2uhtgv2j4"
}
```

- **ID de cliente de OAuth:** Solo descargamos y guardamos el archivo, nombra al archivo como mejor convengas, el contenido será algo como esto:

```json
{
  "web": {
    "client_id": "293831635874-8dfdmnbctsmfhsgfhg874.apps.googleusercontent.com",
    "project_id": "elegant-tangent-388222",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_secret": "GOCSPX-...",
    "redirect_uris": ["http://localhost:5500/"],
    "javascript_origins": ["http://localhost:5500"]
  }
}
```

- **Token de acceso:** En este apartado mencioné que solo serián dos archivos, hay un tercer archivo en formato json, este archivo será generado por la instancia de `FlowOauth` cuando lo invoques desde **VBA** al intentar acceder a tu **Google Drive**. Solo debes asegurarte de pasarlo en el argumento `credentialsPathToken` la ruta de dicho archivo a la instancia de `FlowOauth`. Nombra al archivo como mejor convengas.

> [!NOTE]
> En ningún caso será necesario crear el archivo con el **token de acceso** de forma manual, la instancia de `FlowOauth` se encargará de crearlo si no lo encuentra o actualizarlo según corresponda.

## Probar FlowOauth y generar el token de acceso

Copia el código de abajo en un módulo de **VBA** y ejecutalo eso deberá devolver un **token de acceso** y una **clave de API**.
- La primera vez que ejecutes este código,sigue estos paso:

  1. Selecciona tu cuenta google.
  2. Luego se mostrará una ventana **Google no ha verificado esta aplicación**; selecciona la opción de **continuar**.
  3. Se te mostrará un ventana indicandote los permisos que estas otorgando para acceder a tus **Google Drive**, selecciona la opción de **continuar**.
  4. La siguiente vista será un **No se puede encontrar esta página (localhost)**, debes ir a la barra de direcciones y copiar el valor de `code`(la parte que indica **code=**`copiar_valor`**&scope**), habrás notado que hay cuadro de diálogo **inputbox** esperando que pegues ese valor, después de aceptar se habrá generado el token en la ruta que le has indicado.

```vb
Sub test_oauth()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    Dim oauth20 As New GoogleApis.FlowOauth
    Dim scopes As String
 
    credentialsClient = ThisWorkbook.Path + "\credentials\clientweb.json"
    credentialsToken = ThisWorkbook.Path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.Path + "\credentials\apikey.json"
    scopes = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheet"
    
    With oauth20
        .InitializeFlow credentialsClient, credentialsToken, credentialsApikey, scopes
        Debug.Print .accessToken
        Debug.Print .apiKey
    End With
    
End Sub
```


## Ejemplo de uso

### Listar archivos de Google Drive

```vb
Function getCredentials() As GoogleApis.FlowOauth
     
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    Dim oauth20 As New GoogleApis.FlowOauth
    Dim scopes As String
     
    credentialsClient = ThisWorkbook.Path + "\credentials\clientweb.json"
    credentialsToken = ThisWorkbook.Path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.Path + "\credentials\apikey.json"
    scopes = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheet"
    
    With oauth20
        .InitializeFlow credentialsClient, credentialsToken, credentialsApikey, scopes
    End With
    
    Set getCredentials = oauth20
    
End Function

Sub list_file()
    
    Dim drive As New GoogleApis.GoogleDrive
    Dim result As String
    
    On Error GoTo Catch
    
    With drive
        .connectionService getCredentials
        
        With .files
            result = .list()
            Debug.Print result
            
        End With
    End With
    
    Exit Sub
    
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
    
End Sub
```
En la **ventana inmediato** tendremos una salida como esta:

```json
{
  "files": [
    {
      "mimeType": "application/vnd.google-apps.folder",
      "id": "1_ueAj6ORZjEIGAh7mdGkR_W5JNWoyhdy85",
      "name": "Proyecto salmon"
    },
    {
      "mimeType": "application/vnd.google-apps.folder",
      "id": "1EIX-exARi3UxhT1FNac75HvUJ0aYUHy2",
      "name": "New Folder"
    }
  ]
}
```

> [!NOTE]
> La salida es referencial, devolverá una lista de archivos.



## Recursos adicionales
Los siguientes enlaces estan relacionados a las consultas para listar.

- [Method: files.list](https://developers.google.com/drive/api/v3/reference/files/list)
- [Buscar carpetas o archivos específicos en la sección Mi unidad del usuario actual](https://developers.google.com/drive/api/guides/search-files?hl=es-419#specific)
- [Ejemplos de cadenas de consulta](https://developers.google.com/drive/api/guides/search-files?hl=es-419#examples)