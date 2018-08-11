Imports System.Globalization
Imports System.Data.OleDb
Imports DevExpress.XtraSplashScreen

Module Module1

    Public userprog, pwd As String

    Public Function convert_date_to_eng(ByVal valdate As String) As String

        If valdate = "" Then
            Return ""
        End If

        valdate = CType(valdate, DateTime).ToString("MM/dd/yyyy", CultureInfo.CreateSpecificCulture("en-US"))

        Return valdate

    End Function

    Public Function convert_date_to_ind(ByVal valdate As String) As String

        If valdate = "" Then
            Return ""
        End If

        valdate = CType(valdate, Date).ToString("dd/MM/yyyy", CultureInfo.CreateSpecificCulture("id-ID"))

        Return valdate

    End Function

    Public Sub open_wait()
        SplashScreenManager.ShowForm(Form1, GetType(waitf), True, True, False)
    End Sub

    Public Sub close_wait()
        SplashScreenManager.CloseForm(False)
    End Sub

End Module
