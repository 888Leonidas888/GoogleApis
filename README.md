﻿# ☝️Google APIs para VBA

![VBA](https://img.shields.io/badge/VBA-4c9c4c?style=flat&logo=microsoft&logoColor=white)
![API Google Drive](https://img.shields.io/badge/Api%20Google%20Drive-v3-cccccc?style=flat&logo=googledrive&logoColor=white)
![API Google Sheets](https://img.shields.io/badge/Api%20Google%20Sheets-v4-441133?style=flat&logo=googlesheets&logoColor=white)
![OAuth 2.0](https://img.shields.io/badge/OAuth%202.0-Authorization-blue?style=flat&logo=auth0&logoColor=white)


🚀Consuma las APIs de Google Drive y Sheets desde **VBA** con esta DLL escrita en **C#** para **VBA**. 

Eh tratado de abarcar la mayor cantidad de operaciones que le permiten cada API,
desde listar archivos, cargar archivos, descargar archivos, crear carpetas, eliminar archivos, etc. Y para google sheets, leer y escribir en hojas de cálculo, 
crear hojas de cálculo, eliminar hojas de cálculo, actualizar hojas de cálculo,etc.

📑Esta DLL también provee una clase llamada `FlowOauth` que se encarga de gestionar el flujo de autenticación y autorización de **OAuth 2.0** para acceder a los recursos de Google Drive y Google Sheets, por lo que no tendrás que preocuparte por la autenticación y autorización de tus aplicaciones.

Los métodos expuestos se basan en la documentación oficial de Google, por lo que si deseas hacer una consulta más específica, te recomiendo que visites los siguientes enlaces:

- [Google Drive API](https://developers.google.com/drive/api/v3/reference)
- [Google Sheets API](https://developers.google.com/sheets/api/reference/rest)
- [Oauth 2.0](https://developers.google.com/identity/protocols/oauth2)


> [!NOTE]
> Dentro de la documentación oficial de Google, encontrarás un apartado donde probar cada método, te recomiendo que pases por ahí antes de echar código.

## Hay mucho por desarrollar

🔥Eh tratado de cubrir la mayor parte de los métodos, sin embargo aún queda mucho por cubrir, si quieres colaborar con el desarrollo de este repositorio haz un fork y sube una PR para discutirlo.

> [!NOTE]
> La función de carga de recursos cubren **carga simple**, **carga multiparte** recomendada para la carga `<= 5mb` revisa para mas detalles [subir datos de archivos](https://developers.google.com/drive/api/guides/manage-uploads?hl=es-419). Cargas mayores no estan soportadas por ahora.

## Tabla de contenido
+ 1️⃣ [Instalación](#instalación)
+ 2️⃣ [Activar referencias](#activar-referencias)
+ 3️⃣ [Configuración de entorno en Google](#Configuración-de-entorno-en-google)
+ 4️⃣ [Guardar credenciales de acceso](#guardar-credenciales-de-acceso)
+ 5️⃣ [Probar FlowOauth y generar el token de acceso](#probar-flowoauth-y-generar-el-token-de-acceso)
+ 6️⃣ [Ejemplo de uso](#ejemplo-de-uso)
+ 7️⃣ [Video tutorial](#video-tutorial)
+ 8️⃣ [Recursos adicionales](#Recursos-adicionales)

## 1️⃣ Instalación

👇 Descargue, descomprima y ejecute el archivo **setup.exe** para instalar la DLL.


[![Descargar Instalador](https://img.shields.io/badge/⬇-Descargar%207z-green?style=for-the-badge)](https://github.com/888Leonidas888/GoogleApis/releases/download/v1.0.0/setup.7z)


## 2️⃣ Activar referencias

Antes de hacer uso, debes asegurarte de tener activadas las siguientes referencias, una vez abierto el archivo desde cual usarás esta DLL te 
saltará una advertencia pidiendo que actives las macros, acepta para continuar, una vez habilitada las macros 
presiona `Alt` + `F11` para ir al VBE. En la barra de menú seleciona **Herramientas** -> **Referencias** -> 
**Examinar**, busca la carpeta donde descargaste y selecciona el siguiente archivo **GoogleApis.tlb**, deberás tener marcada estas referencias:

+ ✅ GoogleApis
+ ✅ Microsoft Scripting Runtime



> [!NOTE]
> Aparte de las referencias mencionadas líneas arriba, también se debe contar con el siguiente módulo 
[JsonConverter.bas v2.3.1](https://github.com/VBA-tools/VBA-JSON/tree/master), este módulo te facilitará 
la lectura y escritura de archivos **.json**.



## 3️⃣ Configuración de entorno en google

Posiblemente este sea uno de los pasos mas tediosos a seguir pero tomese su tiempo para leerlo detenidamente, visite la documentación para saber mas [Desarrolla en Google Workspace](https://developers.google.com/workspace/guides/get-started?hl=es_419).

Aquí te dejo un video de como hacerlo, lo hize hace un algún tiempo pero el proceso sigue siendo el mismo, debo mencionar que para agregar mas apis a tu proyecto
debes seguir el paso del minuto 1:08 que esta en el video, por ahora solo podemos agregar la **API de Google Drive** y **Google Sheets**.

📺 [Ver en Youtube](https://youtu.be/iW2cC6AZ7i8)

[![Mira el video](./Assets/Captura_de_pantalla_2025-02-15_104855.png)](https://youtu.be/8GG7LnaMtuE?si=JXTi411BBu7Hsbn1)





1. [Crea un proyecto de Google Cloud](https://developers.google.com/workspace/guides/create-project?hl=es-419)
2. [Habilita las APIs que deseas usar](https://developers.google.com/workspace/guides/enable-apis?hl=es-419)
3. [Obtén información sobre cómo funcionan la autenticación y autorización](https://developers.google.com/workspace/guides/auth-overview?hl=es-419)
4. [Configura el consentimiento de OAuth](https://developers.google.com/workspace/guides/configure-oauth-consent?hl=es-419)
5. [Crea credenciales de acceso](https://developers.google.com/workspace/guides/create-credentials?hl=es-419)

## 4️⃣ Guardar credenciales de acceso

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

## 5️⃣ Probar FlowOauth y generar el token de acceso

👨‍💻 Copia el código de abajo en un módulo de **VBA** y ejecutalo eso deberá devolver un **token de acceso** y una **clave de API**.
- La primera vez que ejecutes este código,sigue estos paso:

  1. Selecciona tu cuenta google.
  2. Luego se mostrará una ventana **Google no ha verificado esta aplicación**; selecciona la opción de **continuar**.
  3. Se te mostrará un ventana indicandote los permisos que estas otorgando para acceder a tus **Google Drive**, selecciona la opción de **continuar**.
  4. Si todo sale bien, se te mostrará un mensaje de **Código de autorización recibido. Puedes cerrar esta ventana.**.

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


## 6️⃣ Ejemplo de uso

En las siguientes líneas de código se muestra un ejemplo de como listar archivos en tu **Google Drive**,
crear una carpeta, mover un archivo a una carpeta, crear una hoja de cálculo y enviar datos a una hoja de cálculo.

Las clases de **GoogleDrive** y **GoogleSheets** tiene mas por ofrecerte, esto es solo un ejemplo de como usarlas.

```vb
Public Const OU_SCOPE_DRIVE  As String = "https://www.googleapis.com/auth/drive"
Public Const OU_SCOPE_DRIVE_FILE As String = "https://www.googleapis.com/auth/drive.file"
Public Const OU_SCOPE_DRIVE_METADATA_READONLY As String = "https://www.googleapis.com/auth/drive.metadata.readonly"
Public Const OU_SCOPE_DRIVE_APPDATA As String = "https://www.googleapis.com/auth/drive.appdata"
Public Const OU_SCOPE_DRIVE_METADATA As String = "https://www.googleapis.com/auth/drive.metadata"
Public Const OU_SCOPE_SPREADSHEETS As String = "https://www.googleapis.com/auth/spreadsheets"
Public Const OU_SCOPE_DRIVE_READONLY As String = "https://www.googleapis.com/auth/drive.readonly"
Public Const OU_SCOPE_SPREADSHEETS_READONLY As String = "https://www.googleapis.com/auth/spreadsheets.readonly"
Public Const OU_SCOPE_PHOTOS_READONLY As String = "https://www.googleapis.com/auth/drive.photos.readonly"
Public Const OU_SCOPE_DRIVE_SCRIPTS As String = "https://www.googleapis.com/auth/drive.scripts"

Function getCredentials() As GoogleApis.FlowOauth
     
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    Dim oauth20 As New GoogleApis.FlowOauth
    Dim scopes As String
     
    credentialsClient = ThisWorkbook.Path + "\credentials\clientweb.json"
    credentialsToken = ThisWorkbook.Path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.Path + "\credentials\apikey.json"
    scopes = OU_SCOPE_DRIVE & " " & OU_SCOPE_SPREADSHEETS
    
    With oauth20
        .InitializeFlow credentialsClient, credentialsToken, credentialsApikey, scopes
    End With
    
    Set getCredentials = oauth20
    
End Function

Sub list_file()
    
    Dim drive As New GoogleApis.GoogleDrive
    Dim queryParameters As New Dictionary
    Dim result As String
    
    On Error GoTo Catch
    
    With queryParameters
        .Add "q", "mimeType = 'application/vnd.google-apps.folder'"
    End With
    
    With drive
        .ConnectionService getCredentials
        
        With .Files
            result = .List(queryParameters)
            Debug.Print result
            
        End With
    End With
    
    Exit Sub
    
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
    
End Sub

Sub create_folder()

    Dim drive As New GoogleApis.GoogleDrive
    Dim result As String
    
    On Error GoTo Catch
    
    With drive
        .ConnectionService getCredentials
        
        With .Files
            result = .UploadMedia()
            Debug.Print result
            
        End With
    End With
    
    Exit Sub
    
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
    
End Sub

Sub create_sheet()
    
    Dim sh As New GoogleApis.GoogleSheets
    
    With sh
        .ConnectionService getCredentials
        Debug.Print .SpreadSheets.Create()
    End With
    
End Sub

Sub move_sheet_to_folder()
        
    Dim drive As New GoogleApis.GoogleDrive
    Dim queryParameters As New Dictionary
    Dim result As String
    Dim fileID As String

    On Error GoTo Catch
    
    With queryParameters
        .Add "addParents", "1WKcZluaXv0mFmc8OyNNBN8wCqreH4vMo"
    End With
    
    fileID = "1yi3l-CQefxRe40jzbFSCUO8MbJ4DtF1mzkfi3e9Q8gs"
     
    With drive
        .ConnectionService getCredentials
        
        With .Files
            result = .Update(fileID, queryParameters:=queryParameters)
            Debug.Print result
            
        End With
    End With
    
    Exit Sub
    
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
End Sub


Sub send_data_sheet()
    
    Dim sh As New GoogleApis.GoogleSheets
    Dim rsl As String
    Dim spreadSheetID$, rng$, json$
    Dim queryParameters As New Dictionary
    Dim fileObject As New Dictionary
    Dim vl As New Collection
    Dim title As New Collection
    Dim reg1 As New Collection
    Dim reg2 As New Collection
    Dim reg3 As New Collection
    
    queryParameters.Add "valueInputOption", "USER_ENTERED"
    spreadSheetID = "1yi3l-CQefxRe40jzbFSCUO8MbJ4DtF1mzkfi3e9Q8gs"
    rng = "Hoja 1"
    
    With title
        .Add "LIBRO"
        .Add "AUTOR"
        .Add "ISBN"
        .Add "PUBLICACION"
    End With

    With reg1
        .Add "=upper(""aplicaciones vba con excel"")"
        .Add "manuel torres remon"
        .Add "978-612-304-2653"
        .Add "2015"
    End With
    
    With reg2
        .Add "=upper(""curso javascript"")"
        .Add "astror de caso parra"
        .Add "978-84-415-42280"
        .Add "=now()"
    End With
    
    With reg3
        .Add "=upper(""c# 9 desarrolle aplicaciones con visual studio 2019"")"
        .Add "jerome hugon"
        .Add "978-2-409-03264-6"
        .Add "2021"
    End With
    
    With vl
        .Add title
        .Add reg1
        .Add reg2
        .Add reg3
    End With
    
    
    With fileObject
        .Add "range", "Hoja 1"
        .Add "values", vl
    End With
    
    json = JsonConverter.ConvertToJson(fileObject, 2)

    With sh
        .ConnectionService getCredentials
        With .SpreadSheets
            With .Values
                
                rsl = .Append(spreadSheetID, rng, json, queryParameters)
                
                If .Operation = 200 Then
                    Debug.Print rsl
                    Debug.Print .Operation
                    Debug.Print "ingreso exito"
                Else
                    Debug.Print "Falla al ingresar"
                    Debug.Print rsl
                End If
                
            End With
        End With
    End With
    
End Sub

```
## 7️⃣ Video tutorial ![YouTube](https://img.shields.io/badge/YouTube-Subscribe-F00?style=flat&logo=youtube&logoColor=white)

+ 📺 [Crear credencial es la plataforma de Google](https://youtu.be/iW2cC6AZ7i8)
+ 📺 [Instalación de DLL + demo](https://youtu.be/GTSw4W14IAU)


## 8️⃣ Recursos adicionales
Los siguientes enlaces estan relacionados a las consultas para listar.

- [Method: files.list](https://developers.google.com/drive/api/v3/reference/files/list)
- [Buscar carpetas o archivos específicos en la sección Mi unidad del usuario actual](https://developers.google.com/drive/api/guides/search-files?hl=es-419#specific)
- [Ejemplos de cadenas de consulta](https://developers.google.com/drive/api/guides/search-files?hl=es-419#examples)