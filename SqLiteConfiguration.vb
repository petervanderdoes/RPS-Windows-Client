

<System.Runtime.InteropServices.ComVisible(False)>
Public Class SqLiteConfiguration
    Inherits Entity.DbConfiguration

    Public Sub New()
        SetProviderFactory("System.Data.SQLite", SQLite.SQLiteFactory.Instance)
        SetProviderFactory("System.Data.SQLite.EF6", System.Data.SQLite.EF6.SQLiteProviderFactory.Instance)
        SetProviderServices("System.Data.SQLite",
                            DirectCast(
                                System.Data.SQLite.EF6.SQLiteProviderFactory.Instance.GetService(
                                    GetType(System.Data.Entity.Core.Common.DbProviderServices)),
                                System.Data.Entity.Core.Common.DbProviderServices))
    End Sub
End Class
