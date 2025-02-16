Attribute VB_Name = "demo"
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


