Attribute VB_Name = "Tables"
'@Lang VBA


Option Compare Database
Option Explicit

Public Function Users(Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    Set Users = DbContext.Table(DbPool.DbUser, "Users", ReadOnly)
End Function
