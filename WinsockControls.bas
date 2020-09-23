Attribute VB_Name = "WinsockControls"
'Declarations
'------------

'---------------------------------
'Used with the SplitProxy Function
'---------------------------------
Public Enum ProxyArray
    Proxy = 1
    Port = 2
End Enum

'-------------------------
'Get the server from a URL
'-------------------------
Public Function GetServer(URL As String) As String

    Dim NewURL As String
    NewURL = URL
    
    'checking for the format of the URL
    If InStr(NewURL, "://") Then _
        NewURL = Mid(NewURL, InStr(NewURL, "://") + 3) 'get rid of un-needed info.
        
    If InStr(NewURL, "www.") Then _
        NewURL = Mid(NewURL, InStr(NewURL, "www.") + 4) 'get rid of un-needed info.
        
    If InStr(NewURL, "/") Then _
        NewURL = Left(NewURL, InStr(NewURL, "/") - 1) 'get rid of un-needed info.
        
    'display the server
    If NewURL <> "" Then
        
        GetServer = NewURL
        
    Else
        
        'in case of error
        GetServer = "Error Getting Server."
        
    End If
    
End Function

'----------------------------------------------------
'This is the Header used to GET html data given a URL
'----------------------------------------------------
Public Function SendHeader(URL As String) As String

    Dim NewHeader As String
    
    NewHeader = "GET " & URL & " HTTP/1.0" & vbCrLf
    NewHeader = NewHeader & "Connection: Keep-Alive" & vbCrLf
    NewHeader = NewHeader & "Accept: */*" & vbCrLf
    NewHeader = NewHeader & "Accept-Language: en" & vbCrLf & vbCrLf
    
    SendHeader = NewHeader
    
End Function

'--------------------------------------------------------
'If the page returns a cookie this will 'grab' the cookie
'--------------------------------------------------------
Public Function GrabCookie(Data As String) As String

    Dim NewData As String
    NewData = Data
    
    'check if there is a cooke
    If InStr(Data, "Set-Cookie") Then
        
        'keep adding the cookies until there is no more
        While InStr(NewData, "Set-Cookie")
            
            'parse the cookie
            NewData = Mid(NewData, InStr(NewData, "Set-Cookie") + 12)
            NewData = Left(NewData, InStr(NewData, ";"))
            GrabCookie = GrabCookie & NewData & " "
            
        Wend
        
    Else
    
        'no cookie
        GrabCookie = "No Cookie Available."
        
    End If
    
End Function

'-----------------------------------
'Fixes the linefeed on returned html
'-----------------------------------
Public Function FixFeed(Data As String) As String

    Dim NewData As String
    'fix line feeds that are already correct
    NewData = Replace(Data, Chr(13) & Chr(10), "BEGIN:::REPLACE THIS AFTER:::END")
    
    'fix bad line feeds and hide
    NewData = Replace(NewData, Chr(10), Chr(13) & Chr(10))
    NewData = Replace(NewData, Chr(13) & Chr(10), "BEGIN:::REPLACE THIS AFTER:::END")
    
    'fix bad other bad line feeds
    NewData = Replace(NewData, Chr(13), Chr(13) & Chr(10))
    
    'unhide corrected linefeeds
    NewData = Replace(NewData, "BEGIN:::REPLACE THIS AFTER:::END", Chr(13) & Chr(10))
    
    FixFeed = NewData
    
End Function

'--------------
'Splits a proxy
'--------------
Public Function SplitProxy(Proxy As String, Delimiter As String, Output As ProxyArray) As String
    If InStr(Proxy, Delimiter) Then
        If Output = 1 Then
            SplitProxy = Trim(Left(Proxy, InStr(Proxy, Delimiter) - 1))
        Else
            SplitProxy = Trim(Mid(Proxy, InStr(Proxy, Delimiter) + 1))
        End If
    Else
        SplitProxy = "Error"
    End If
End Function
