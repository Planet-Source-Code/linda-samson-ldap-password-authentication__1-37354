<div align="center">

## LDAP Password Authentication


</div>

### Description

Show how to use LDAP to authenticate users
 
### More Info
 
LDAP Server, LDAP Port, UID, Organization, Organizational Unit, Password

Requires Microsoft ActiveX Data Objects 2.0/2.1/2.5 Library

This is just a code snippet. You must supply the command button, etc.

True - if authenticated, False - if not


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[linda samson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/linda-samson.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/linda-samson-ldap-password-authentication__1-37354/archive/master.zip)





### Source Code

```
'linda
'Requires Microsoft ActiveX Data Objects 2.0/2.1/2.5 Library (Project | References)
'TODO:
'Private Sub Command1_Click()
'  Init()
'  msgbox Authenticate
'End Sub
dim m_LDAPServer as string
dim m_LDAPPort as string
dim m_Org as string
dim m_OrgUnit as string
dim m_Initial as string
dim m_Password as string
Public Sub Init()
  m_LDAPServer = "ldapserver.com"
  m_LDAPPort  = "389"
  m_Org    = "ldaporg.com"	'o=
  m_OrgUnit  = "People"       'ou=
  m_Initial  = "userId"       'uid=
  m_Password  = "password"
End Sub
Public Function Authenticate() As Boolean
On Error Resume Next
Dim con As New Connection
Dim sqlStmt As String
Dim connStr As String
Dim rs As Recordset
  'TODO: trap blank UID and password here!
  If (m_Initial = "") then
    Msgbox "No Initial"
    Exit Function
  End If
  if (m_Password = "") then
    Msgbox "No Password"
    Exit Function
  End If
  'prepare SQL statement
  sqlStmt = "SELECT uid " & _
     "FROM 'LDAP://" & m_LDAPServer & ":" & m_LDAPPort & "/o=" & m_Org & "/ou=" & m_OrgUnit & "' " & _
     "WHERE uid='" & m_Initial & "'" & " and objectClass='*'"
  'create Active Directory Service Object
  Set con = CreateObject("adodb.connection")
  con.Provider = "ADSDSOOBject"
  'construct connection string
  connStr = "uid=" & m_Initial & ",ou=" & m_OrgUnit & ",o=" & m_Org
  'open connection + password
  con.Open "ADs Provider", connStr, m_Password
  'execute our query
  Set rs = con.Execute(sqlStmt)
  'rs will be empty if authentication fails
  Authenticate = Not (IsEmpty(rs) Or (Err.Number = -2147217911))
  'need to close
  rs.Close
End Function
```

