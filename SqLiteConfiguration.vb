Imports System.Data.Entity
Imports System.Data.Entity.Core.Common
Imports System.Data.SQLite
Imports System.Data.SQLite.EF6
Imports System.Runtime.InteropServices

<ComVisible(False)>
Public Class SqLiteConfiguration
    Inherits DbConfiguration

    Public Sub New()
        SetProviderFactory("System.Data.SQLite", SQLiteFactory.Instance)
        SetProviderFactory("System.Data.SQLite.EF6", SQLiteProviderFactory.Instance)
        SetProviderServices("System.Data.SQLite",
                            DirectCast(SQLiteProviderFactory.Instance.GetService(GetType(DbProviderServices)),
                                       DbProviderServices))
    End Sub
End Class
