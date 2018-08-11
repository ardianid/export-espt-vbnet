Imports System.Data
Imports System.Data.OleDb
Imports DevExpress.XtraGrid.Views.Grid
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Imports DevExpress.XtraEditors

Public Class Form1

    Private dvmanager1 As Data.DataViewManager
    Private dv1 As Data.DataView

    Private dvmanager2 As Data.DataViewManager
    Private dv2 As Data.DataView

    Private dvmanager3 As Data.DataViewManager
    Private dv3 As Data.DataView

    Private dvmanager4 As Data.DataViewManager
    Private dv4 As Data.DataView

    Private dvmanager_det As Data.DataViewManager
    Private dv_det As Data.DataView

    Private dvmanager_ret As Data.DataViewManager
    Private dv_ret As Data.DataView

    Dim crtotal_perbulan As Rpenjualan
    Dim crtotal_pembelian As Rpembelian

    Private Sub load_data_pkp()

        Dim sql As String = String.Format("select * from ms_cust")

        Dim cn As OleDbConnection = Nothing
        Dim ds As DataSet

        grid1.DataSource = Nothing

        Try

            dv1 = Nothing

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            ds = New DataSet()
            ds = ClassMy.GetDataSet(sql, cn)

            dvmanager1 = New DataViewManager(ds)
            dv1 = dvmanager1.CreateDataView(ds.Tables(0))

            grid1.DataSource = dv1

        Catch ex As OleDb.OleDbException
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally


            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If

        End Try

    End Sub

    Private Sub load_cust_mobiz()

        grid2.DataSource = Nothing

        If IsNothing(dv1) Then
            Return
        End If

        If dv1.Count <= 0 Then
            Return
        End If

        Dim sql As String = String.Format("select custp.KD_PAJAK,custp.KD_PROG as kd_prog,cust.Name as nama,cust2.Address as alamat from ms_cust2 custp " & _
        "inner join M1Company_Live.Shared.BusinessPartners cust " & _
        "on custp.KD_PROG=cust.BusinessPartnerId " & _
        "inner join M1Company_Live.Shared.BusinessAddresses cust2 " & _
        "on custp.KD_PROG=cust2.BusinessPartnerId " & _
        "where custp.KD_PAJAK='{0}'", dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString)

        Dim cn As OleDbConnection = Nothing
        Dim ds As DataSet

        Try

            dv2 = Nothing

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            ds = New DataSet()
            ds = ClassMy.GetDataSet(sql, cn)

            dvmanager2 = New DataViewManager(ds)
            dv2 = dvmanager2.CreateDataView(ds.Tables(0))

            grid2.DataSource = dv2

        Catch ex As OleDb.OleDbException
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If

        End Try

    End Sub

    Private Sub r_cust_pkp()

        Dim cn As OleDbConnection = Nothing
        Try

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            Dim sql As String = "select KD_PAJAK,NAMA from ms_cust where SAKTIF=1"

            Dim ds As DataSet = New DataSet
            ds = ClassMy.GetDataSet(sql, cn)

            tpelanggan.Properties.DataSource = ds.Tables(0)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally


            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If
        End Try

    End Sub

    Private Sub load_data()

        Dim sql As String = String.Format("select dh.InvoiceNo as nobukti,convert(date,dh.TransactionDate) as tanggal,cust.NAMA as nama_pkp,cust.ALAMAT as alamat_pkp,cust.NPWP as npwp_pkp,part.Name as nama_mobiz " & _
        "from ms_cust cust inner join ms_cust2 cust2 " & _
        "on cust.KD_PAJAK=cust2.KD_PAJAK " & _
        "inner join M1Company_Live.Sales.DirectInvoiceHeader dh " & _
        "on cust2.KD_PROG=dh.CustomerId " & _
        "inner join M1Company_Live.Shared.BusinessPartners part " & _
        " on part.BusinessPartnerId=dh.CustomerId " & _
        "where cust.SAKTIF = 1 And dh.Status = 3 " & _
        "and dh.InvoiceNo in (select InvoiceNo from v_fakturcount_item " & _
        "where not(jml_gln>0 and jml_ngln=0)) " & _
        "and cust.KD_PAJAK='{0}' " & _
        "and convert(date,dh.TransactionDate)>='{1}' and convert(date,dh.TransactionDate)<='{2}'", tpelanggan.EditValue, convert_date_to_eng(ttgl1.EditValue), convert_date_to_eng(ttgl2.EditValue))

        Dim cn As OleDbConnection = Nothing
        Dim ds As DataSet

        grid3.DataSource = Nothing

        Try

            dv3 = Nothing

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            ds = New DataSet()
            ds = ClassMy.GetDataSet(sql, cn)

            dvmanager3 = New DataViewManager(ds)
            dv3 = dvmanager3.CreateDataView(ds.Tables(0))

            grid3.DataSource = dv3

            tseri.EditValue = dv3.Count

        Catch ex As OleDb.OleDbException
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally


            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If

        End Try

    End Sub

    Private Sub load_noseri()

        Dim sql As String = String.Format("select * from temptable")

        Dim cn As OleDbConnection = Nothing
        Dim ds As DataSet

        grid4.DataSource = Nothing

        Try

            dv4 = Nothing

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            ds = New DataSet()
            ds = ClassMy.GetDataSet(sql, cn)

            dvmanager4 = New DataViewManager(ds)
            dv4 = dvmanager4.CreateDataView(ds.Tables(0))

            grid4.DataSource = dv4

        Catch ex As OleDb.OleDbException
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally


            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If

        End Try

    End Sub

    Private Sub cek_indo_alfa_baru()

        Dim cn As OleDbConnection = Nothing

        Try

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            Dim Sql As String = "select a.BusinessPartnerId,a.Name from M1Company_Live.Shared.BusinessPartners a where a.IsActive=1 and " & _
            "(a.Name like 'INDOMARET%' OR a.Name like 'ALFAMART%') " & _
            "and  not(a.BusinessPartnerId in  (SELECT KD_PROG from db_penghubung.dbo.ms_cust2 " & _
            "where KD_PAJAK in ('ALFA','INDO'))) " & _
            "and a.BusinessPartnerId in (select CustomerId from M1Company_Live.sales.DirectInvoiceHeader where Status=3)"

            Dim cmd As OleDbCommand = New OleDbCommand(Sql, cn)
            Dim drd As OleDbDataReader = cmd.ExecuteReader

            If drd.HasRows Then
                While drd.Read

                    Dim kode As String = drd(0).ToString
                    Dim nama As String = drd(1).ToString
                    nama = nama.Trim.Substring(0, 4)

                    Dim sqlins As String = String.Format("insert into db_penghubung.dbo.ms_cust2 (KD_PAJAK,KD_PROG) values('{0}','{1}')", nama, kode)
                    Using cmdin As OleDbCommand = New OleDbCommand(sqlins, cn)
                        cmdin.ExecuteReader()
                    End Using

                End While
            End If


        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If
        End Try

    End Sub

    Private Sub export_data()

        If IsNothing(dv4) Then
            MsgBox("No seri harus diisi")
            Return
        End If

        If dv4.Count <= 0 Then
            MsgBox("No seri harus diisi")
            Return
        End If

        Dim totseri As Integer = 0
        For i As Integer = 0 To dv4.Count - 1
            totseri = totseri + dv4(i)("norange").ToString
        Next

        If totseri < tseri.EditValue Then
            MsgBox("No seri tidak mencukupi untuk export data ini, export dibatalkan", vbOKOnly + vbExclamation, "Informasi")
            Return
        End If

        open_wait()

        Dim xlAppToUpload As New Excel.Application
        xlAppToUpload.Workbooks.Add()

        Dim xlWorkSheetToUpload As Excel.Worksheet
        xlWorkSheetToUpload = xlAppToUpload.Sheets("Sheet1")


        ' cells (row, column)

        With xlWorkSheetToUpload
            .Cells(1, 1).value = "FK"
            .Cells(1, 2).value = "KD_JENIS_TRANSAKSI"
            .Cells(1, 3).value = "FG_PENGGANTI"
            .Cells(1, 4).value = "NOMOR_FAKTUR"
            .Cells(1, 5).value = "MASA_PAJAK"
            .Cells(1, 6).value = "TAHUN_PAJAK"
            .Cells(1, 7).value = "TANGGAL_FAKTUR"
            .Cells(1, 8).value = "NPWP"
            .Cells(1, 9).value = "NAMA"
            .Cells(1, 10).value = "ALAMAT_LENGKAP"
            .Cells(1, 11).value = "JUMLAH_DPP"
            .Cells(1, 12).value = "JUMLAH_PPN"
            .Cells(1, 13).value = "JUMLAH_PPNBM"
            .Cells(1, 14).value = "ID_KETERANGAN_TAMBAHAN"
            .Cells(1, 15).value = "FG_UANG_MUKA"
            .Cells(1, 16).value = "UANG_MUKA_DPP"
            .Cells(1, 17).value = "UANG_MUKA_PPN"
            .Cells(1, 18).value = "UANG_MUKA_PPNBM"
            .Cells(1, 19).value = "REFERENSI"

            .Cells(2, 1).value = "LT"
            .Cells(2, 2).value = "NPWP"
            .Cells(2, 3).value = "NAMA"
            .Cells(2, 4).value = "JALAN"
            .Cells(2, 5).value = "BLOK"
            .Cells(2, 6).value = "NOMOR"
            .Cells(2, 7).value = "RT"
            .Cells(2, 8).value = "RW"
            .Cells(2, 9).value = "KECAMATAN"
            .Cells(2, 10).value = "KELURAHAN"
            .Cells(2, 11).value = "KABUPATEN"
            .Cells(2, 12).value = "PROPINSI"
            .Cells(2, 13).value = "KODE_POS"
            .Cells(2, 14).value = "NOMOR_TELEPON"

            .Cells(3, 1).value = "OF"
            .Cells(3, 2).value = "KODE_OBJEK"
            .Cells(3, 3).value = "NAMA"
            .Cells(3, 4).value = "HARGA_SATUAN"
            .Cells(3, 5).value = "JUMLAH_BARANG"
            .Cells(3, 6).value = "HARGA_TOTAL"
            .Cells(3, 7).value = "DISKON"
            .Cells(3, 8).value = "DPP"
            .Cells(3, 9).value = "PPN"
            .Cells(3, 10).value = "TARIF_PPNBM"
            .Cells(3, 11).value = "PPNBM"


        End With

        Dim nobukti As String = ""
        Dim tanggal As String = ""
        Dim npwp As String = ""
        Dim nama_pkp As String = ""
        Dim alamat_pkp As String = ""

        Dim nomor_h As Integer = 3
        Dim nomor_d As Integer = 2

        Dim tot_dpp As Double = 0
        Dim tot_ppn As Double = 0

        Dim hit_dpp As Double = 0
        Dim hit_ppn As Double = 0

        Dim cn As OleDbConnection
        cn = New OleDbConnection
        cn = ClassMy.open_conn_mobiz

        Try

            dv3.Sort = "tanggal asc,nobukti asc"

            Dim norange_seri As Integer = Integer.Parse(dv4(0)("norange").ToString)
            Dim noawal_seri0 As String = dv4(0)("noawal").ToString
            noawal_seri0 = noawal_seri0.Substring(0, noawal_seri0.Length - 8)

            Dim noawal_seri As String = dv4(0)("noawal").ToString
            noawal_seri = Microsoft.VisualBasic.Right(noawal_seri, 8)

            Dim norow_urut As Integer = 0
            Dim norow_seri As Integer = 0
            Dim nomor_totdpp As Integer = 0

            For i As Integer = 0 To dv3.Count - 1

                If norow_urut > norange_seri Then

                    norow_seri = norow_seri + 1
                    norow_urut = 0

                    norange_seri = Integer.Parse(dv4(norow_seri)("norange").ToString)

                    noawal_seri0 = dv4(norow_seri)("noawal").ToString
                    noawal_seri0 = noawal_seri0.Substring(0, noawal_seri0.Length - 8)

                    noawal_seri = dv4(norow_seri)("noawal").ToString
                    noawal_seri = Microsoft.VisualBasic.Right(noawal_seri, 8)

                End If

                nobukti = dv3(i)("nobukti").ToString
                tanggal = convert_date_to_ind(dv3(i)("tanggal").ToString)
                nama_pkp = dv3(i)("nama_pkp").ToString
                alamat_pkp = dv3(i)("alamat_pkp").ToString
                npwp = dv3(i)("npwp_pkp").ToString
                nomor_totdpp = 0

                With xlWorkSheetToUpload
                    .Cells(nomor_h + 1, 1).value = "FK"
                    .Cells(nomor_h + 1, 2).value = "01"

                    .Range("B" & (nomor_h + 1), "B" & (nomor_h + 1)).NumberFormat = "00"

                    .Cells(nomor_h + 1, 3).value = "0"
                    .Cells(nomor_h + 1, 4).value = noawal_seri0 & Double.Parse(noawal_seri) + 1

                    .Range("D" & (nomor_h + 1), "D" & (nomor_h + 1)).NumberFormat = "0000000000000"

                    .Cells(nomor_h + 1, 5).value = tmasa.EditValue
                    .Cells(nomor_h + 1, 6).value = tthn.EditValue
                    .Cells(nomor_h + 1, 7).value = tanggal

                    .Cells(nomor_h + 1, 8).value = npwp
                    .Range("H" & (nomor_h + 1), "H" & (nomor_h + 1)).NumberFormat = "000000000000000"

                    .Cells(nomor_h + 1, 9).value = nama_pkp
                    .Cells(nomor_h + 1, 10).value = alamat_pkp
                    .Cells(nomor_h + 1, 11).value = 0
                    .Cells(nomor_h + 1, 12).value = 0
                    .Cells(nomor_h + 1, 13).value = 0
                    .Cells(nomor_h + 1, 15).value = 0
                    .Cells(nomor_h + 1, 16).value = 0
                    .Cells(nomor_h + 1, 17).value = 0
                    .Cells(nomor_h + 1, 18).value = 0
                    .Cells(nomor_h + 1, 19).value = nobukti

                    If nobukti.ToString.Length = 8 Then
                        .Range("S" & (nomor_h + 1), "S" & (nomor_h + 1)).NumberFormat = "00000000"
                    End If


                    nomor_totdpp = nomor_h + 1

                    .Cells(nomor_h + 2, 1).value = "FAPR"

                    '.Cells(nomor_h + 2, 2).value = "016750812322001"
                    '.Range("B" & (nomor_h + 2), "B" & (nomor_h + 2)).NumberFormat = "000000000000000"

                    .Cells(nomor_h + 2, 2).value = "PT WATERINDEX TIRTA LESTARI"
                    .Cells(nomor_h + 2, 3).value = "JL TEMBESU I NO 1 , BANDAR LAMPUNG"
                    .Cells(nomor_h + 2, 4).value = "Alfian"
                    .Cells(nomor_h + 2, 5).value = "BANDAR LAMPUNG"
                    '.Cells(nomor_h + 2, 7).value = "-"
                    '.Cells(nomor_h + 2, 8).value = "-"
                    '.Cells(nomor_h + 2, 9).value = "-"
                    '.Cells(nomor_h + 2, 10).value = "-"
                    '.Cells(nomor_h + 2, 11).value = "-"
                    '.Cells(nomor_h + 2, 12).value = "-"
                    '.Cells(nomor_h + 2, 13).value = "-"
                    '.Cells(nomor_h + 2, 14).value = "-"

                    nomor_h = nomor_h + 2
                    nomor_d = nomor_h

                    tot_dpp = 0
                    tot_ppn = 0

                    Dim sql As String = String.Format("select ItemId,convert(int,UnitPrice) as harga,convert(int,Quantity) as jml,convert(int,DiscountAmount) as disc from Sales.DirectInvoiceDetail " & _
                    "where not(ItemId='G0003') and Total>0 and InvoiceNo='{0}'", nobukti)
                    Dim cmd As OleDbCommand = New OleDbCommand(sql, cn)
                    Dim drd As OleDbDataReader = cmd.ExecuteReader

                    While drd.Read

                        Dim kd_barang = drd(0).ToString
                        Dim harga As Double = Double.Parse(drd(1).ToString)
                        Dim jml As Integer = Integer.Parse(drd(2).ToString)
                        Dim disc As Double = Double.Parse(drd(3).ToString)

                        Dim nama_barang As String = ""
                        Select Case kd_barang
                            Case "G0001"
                                nama_barang = "AIR GRAND"
                            Case "C0007"
                                nama_barang = "GELAS 150 ML"
                            Case "C0002"
                                nama_barang = "GELAS 240 ML"
                            Case "B0005"
                                nama_barang = "BOTOL 330 ML"
                            Case "B0003"
                                nama_barang = "BOTOL 600 ML"
                            Case "B0002"
                                nama_barang = "BOTOL 1500 ML"
                            Case "B0001"
                                nama_barang = "BOTOL 500 ML"
                            Case "B002A"
                                nama_barang = "BOTOL 1500 ML ALFAMART"
                            Case "B003A"
                                nama_barang = "BOTOL 600 ML ALFAMART"
                            Case "B005A"
                                nama_barang = "BOTOL 330 ML ALFAMART"
                            Case "C002A"
                                nama_barang = "GELAS 240 ML ALFAMART"
                            Case Else
                                nama_barang = "ERRRORRRRRR BARANG TIDAK DITEMUKAN.. (Hub IT)"
                        End Select

                        nomor_d = nomor_d + 1
                        nomor_h = nomor_d

                        hit_dpp = System.Math.Round(((harga - disc) / 1.1), 1)

                        If harga = 9000 Then
                            hit_dpp = hit_dpp + 0.1
                        ElseIf harga = 10000 Then
                            hit_dpp = 9091 - System.Math.Round((disc / 1.1), 1)
                        End If

                        hit_dpp = hit_dpp * jml

                        hit_ppn = hit_dpp * (10 / 100)

                        .Cells(nomor_d, 1).value = "OF"
                        .Cells(nomor_d, 2).value = kd_barang
                        .Cells(nomor_d, 3).value = nama_barang

                        If harga = 9000 Then
                            .Cells(nomor_d, 4).value = System.Math.Round(((harga - disc) / 1.1), 1) + 0.1
                        ElseIf harga = 10000 Then
                            .Cells(nomor_d, 4).value = 9091 - System.Math.Round((disc / 1.1), 1)
                        Else
                            .Cells(nomor_d, 4).value = System.Math.Round(((harga - disc) / 1.1), 1)
                        End If

                        .Cells(nomor_d, 5).value = jml
                        .Cells(nomor_d, 6).value = Math.Floor(hit_dpp)
                        .Cells(nomor_d, 7).value = 0
                        .Cells(nomor_d, 8).value = Math.Floor(hit_dpp)
                        .Cells(nomor_d, 9).value = Math.Floor(hit_ppn)
                        .Cells(nomor_d, 10).value = 0
                        .Cells(nomor_d, 11).value = 0

                        tot_dpp = tot_dpp + Math.Floor(hit_dpp)
                        tot_ppn = tot_ppn + Math.Floor(hit_ppn)

                    End While

                    drd.Close()

                    .Cells(nomor_totdpp, 11).value = Math.Floor(tot_dpp)
                    .Cells(nomor_totdpp, 12).value = Math.Floor(tot_ppn)

                End With

                norow_urut = norow_urut + 1
                noawal_seri = noawal_seri + 1

            Next

            xlAppToUpload.Visible = True
            close_wait()

        Catch ex As Exception
            close_wait()
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If
        End Try

        

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        load_data_pkp()
        load_cust_mobiz()
        load_noseri()

        ttgl1.EditValue = Date.Now
        ttgl2.EditValue = Date.Now

        ttgl1_cr.EditValue = Date.Now
        ttgl2_cr.EditValue = Date.Now

        ttgl1_ret.EditValue = Date.Now
        ttgl2_ret.EditValue = Date.Now

        tbarang_cr.SelectedIndex = 0
        tbarang_ret.SelectedIndex = 0

        tbetul.EditValue = 0
        TextEdit1.EditValue = Year(Date.Now)
        TextEdit2.EditValue = Year(Date.Now)

        cek_indo_alfa_baru()

    End Sub

    Private Sub GridView2_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridView2.FocusedRowChanged
        load_cust_mobiz()
    End Sub

    Private Sub GridView2_RowUpdated(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles GridView2.RowUpdated
        If dv1(Me.BindingContext(dv1).Position)("ID_CUST").ToString = "" Then

            If dv1(Me.BindingContext(dv1).Position)("SAKTIF").ToString = "" Then
                Return
            End If

            Dim nama_cust As String = dv1(Me.BindingContext(dv1).Position)("NAMA").ToString
            Dim alamat_cust As String = dv1(Me.BindingContext(dv1).Position)("ALAMAT").ToString
            Dim kd_cust As String = dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString
            Dim npwp_cust As String = dv1(Me.BindingContext(dv1).Position)("NPWP").ToString
            Dim akt_cust As String = dv1(Me.BindingContext(dv1).Position)("SAKTIF").ToString

            Dim sql As String = String.Format("insert into ms_cust (KD_PAJAK,NAMA,ALAMAT,NPWP,SAKTIF) VALUES('{0}','{1}','{2}','{3}',{4}); Select Scope_Identity()", kd_cust, nama_cust, alamat_cust, npwp_cust, akt_cust)

            Dim cn As OleDbConnection = Nothing

            Try

                cn = New OleDbConnection
                cn = ClassMy.open_conn

                Dim sqltrans As OleDbTransaction = cn.BeginTransaction

                Dim cmd As OleDbCommand
                cmd = New OleDbCommand(sql, cn, sqltrans)

                Dim id As Integer = 0
                id = cmd.ExecuteScalar()

                sqltrans.Commit()

                dv1(Me.BindingContext(dv1).Position)("ID_CUST") = id

            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
            Finally

                If Not cn Is Nothing Then
                    If cn.State = ConnectionState.Open Then
                        cn.Close()
                    End If
                End If
            End Try

        Else

            Dim nama_cust As String = dv1(Me.BindingContext(dv1).Position)("NAMA").ToString
            Dim alamat_cust As String = dv1(Me.BindingContext(dv1).Position)("ALAMAT").ToString
            '' Dim kd_cust As String = dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString
            Dim npwp_cust As String = dv1(Me.BindingContext(dv1).Position)("NPWP").ToString
            Dim akt_cust As String = dv1(Me.BindingContext(dv1).Position)("SAKTIF").ToString

            Dim id_cust As String = dv1(Me.BindingContext(dv1).Position)("ID_CUST").ToString

            Dim sql As String = String.Format("update ms_cust set NAMA='{0}',ALAMAT='{1}',NPWP='{2}',SAKTIF={3} WHERE ID_CUST={4}", nama_cust, alamat_cust, npwp_cust, akt_cust, id_cust)

            Dim cn As OleDbConnection = Nothing

            Try

                cn = New OleDbConnection
                cn = ClassMy.open_conn

                Dim sqltrans As OleDbTransaction = cn.BeginTransaction

                Dim cmd As OleDbCommand
                cmd = New OleDbCommand(sql, cn, sqltrans)

                cmd.ExecuteNonQuery()

                sqltrans.Commit()

            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
            Finally

                If Not cn Is Nothing Then
                    If cn.State = ConnectionState.Open Then
                        cn.Close()
                    End If
                End If
            End Try

        End If
    End Sub

    Private Sub grid1_KeyDown(sender As Object, e As KeyEventArgs) Handles grid1.KeyDown

        If IsNothing(dv1) Then
            Return
        End If

        If dv1.Count <= 0 Then
            Return
        End If

        If e.KeyCode = Keys.Delete Then

            If MsgBox("Akan dihapus ???", vbQuestion + vbYesNo, "Konfirmasi") = MsgBoxResult.Yes Then

                Dim cn As OleDbConnection = Nothing

                Try

                    cn = New OleDbConnection
                    cn = ClassMy.open_conn

                    Dim sqltrans As OleDbTransaction = cn.BeginTransaction

                    Dim sql As String = String.Format("delete from ms_cust where ID_CUST={0}", dv1(Me.BindingContext(dv1).Position)("ID_CUST").ToString)
                    Dim sql2 As String = String.Format("delete from ms_cust2 where KD_PAJAK='{0}'", dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString)

                    Dim cmd As OleDbCommand
                    cmd = New OleDbCommand(sql, cn, sqltrans)

                    cmd.ExecuteNonQuery()

                    Dim cmd2 As OleDbCommand
                    cmd2 = New OleDbCommand(sql2, cn, sqltrans)

                    cmd2.ExecuteNonQuery()

                    sqltrans.Commit()

                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
                Finally

                    If Not cn Is Nothing Then
                        If cn.State = ConnectionState.Open Then
                            cn.Close()
                        End If
                    End If

                    load_data_pkp()

                End Try

            End If


        End If
    End Sub

   
    Private Sub GridView1_RowUpdated(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles GridView1.RowUpdated
        If dv2(Me.BindingContext(dv2).Position)("KD_PAJAK").ToString = "" Then

            Dim nama As String = ""
            Dim alamat As String = ""

            Dim cn As OleDbConnection = Nothing

            Try

                cn = New OleDbConnection
                cn = ClassMy.open_conn_mobiz

                Dim sqlcek As String = String.Format("SELECT KD_PROG from db_penghubung.dbo.ms_cust2 b where b.KD_PAJAK='{0}' and b.KD_PROG='{1}'", dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString, dv2(Me.BindingContext(dv2).Position)("kd_prog").ToString)
                Dim cmdcek As OleDbCommand = New OleDbCommand(sqlcek, cn)
                Dim drdcek As OleDbDataReader = cmdcek.ExecuteReader

                If drdcek.Read Then
                    Dim hasil As String = drdcek(0).ToString

                    If hasil.Trim.Length > 0 Then
                        MsgBox("Customer sudah ada", vbOKOnly + vbInformation, "Informasi")
                        dv2(Me.BindingContext(dv2).Position).Delete()
                        Return
                    End If

                End If
                drdcek.Close()

                Dim sql As String = String.Format("select cust.Name,cust2.Address from M1Company_Live.Shared.BusinessPartners cust " & _
                "inner join M1Company_Live.Shared.BusinessAddresses cust2 " & _
                "on cust.BusinessPartnerId=cust2.BusinessPartnerId where cust.BusinessPartnerId='{0}'", dv2(Me.BindingContext(dv2).Position)("kd_prog").ToString)

                Dim cmd As OleDbCommand = New OleDbCommand(sql, cn)
                Dim drd As OleDbDataReader = cmd.ExecuteReader

                If drd.HasRows Then
                    If drd.Read Then

                        If Not drd(0).ToString.Equals("") Then
                            nama = drd(0).ToString
                            alamat = drd(1).ToString
                        End If

                    End If
                End If
                drd.Close()

                If nama.Trim.Length > 0 Then

                    Dim cn2 As OleDbConnection
                    cn2 = New OleDbConnection
                    cn2 = ClassMy.open_conn

                    Dim sqltrans As OleDbTransaction = cn2.BeginTransaction

                    Dim sql2 As String = String.Format("insert into ms_cust2 (KD_PAJAK,KD_PROG) values('{0}','{1}')", dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString, dv2(Me.BindingContext(dv2).Position)("kd_prog").ToString)

                    Dim cmd2 As OleDbCommand = New OleDbCommand(sql2, cn2, sqltrans)
                    cmd2.ExecuteNonQuery()

                    sqltrans.Commit()

                    dv2(Me.BindingContext(dv2).Position)("KD_PAJAK") = dv1(Me.BindingContext(dv1).Position)("KD_PAJAK").ToString
                    dv2(Me.BindingContext(dv2).Position)("nama") = nama
                    dv2(Me.BindingContext(dv2).Position)("alamat") = alamat

                Else
                    dv2(Me.BindingContext(dv2).Position).Delete()
                    MsgBox("Data tidak ditemukan")
                End If

            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
            Finally

                If Not cn Is Nothing Then
                    If cn.State = ConnectionState.Open Then
                        cn.Close()
                    End If
                End If
            End Try

        End If
    End Sub

    Private Sub grid2_KeyDown(sender As Object, e As KeyEventArgs) Handles grid2.KeyDown

        If IsNothing(dv2) Then
            Return
        End If

        If dv2.Count <= 0 Then
            Return
        End If


        If e.KeyCode = Keys.Delete Then

            If MsgBox("Akan dihapus ???", vbQuestion + vbYesNo, "Konfirmasi") = MsgBoxResult.Yes Then

                Dim cn As OleDbConnection = Nothing

                Try

                    cn = New OleDbConnection
                    cn = ClassMy.open_conn

                    Dim sqltrans As OleDbTransaction = cn.BeginTransaction

                    Dim sql2 As String = String.Format("delete from ms_cust2 where KD_PROG='{0}'", dv2(Me.BindingContext(dv2).Position)("kd_prog").ToString)

                    Dim cmd2 As OleDbCommand
                    cmd2 = New OleDbCommand(sql2, cn, sqltrans)

                    cmd2.ExecuteNonQuery()

                    sqltrans.Commit()

                Catch ex As Exception
                    MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
                Finally

                    If Not cn Is Nothing Then
                        If cn.State = ConnectionState.Open Then
                            cn.Close()
                        End If
                    End If

                    load_cust_mobiz()

                End Try

            End If

        End If
    End Sub

    Private Sub GridView3_RowUpdated(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles GridView3.RowUpdated
        If dv3(Me.BindingContext(dv3).Position)("nama_pkp").ToString = "" Then

            Dim nama As String = ""
            Dim alamat As String = ""
            Dim nama_mobiz As String = ""
            Dim npwp As String = ""

            Dim cn As OleDbConnection = Nothing

            Try

                cn = New OleDbConnection
                cn = ClassMy.open_conn

                Dim sql As String = String.Format("select dh.InvoiceNo as nobukti,convert(date,dh.TransactionDate) as tanggal,cust.NAMA as nama_pkp,cust.ALAMAT as alamat_pkp,cust.NPWP as npwp_pkp,part.Name as nama_mobiz " & _
                "from ms_cust cust inner join ms_cust2 cust2 " & _
                "on cust.KD_PAJAK=cust2.KD_PAJAK " & _
                "inner join M1Company_Live.Sales.DirectInvoiceHeader dh " & _
                "on cust2.KD_PROG=dh.CustomerId " & _
                "inner join M1Company_Live.Sales.DirectInvoiceDetail di " & _
                "on dh.InvoiceNo=di.InvoiceNo " & _
                "inner join M1Company_Live.Shared.BusinessPartners part " & _
                "on part.BusinessPartnerId=dh.CustomerId " & _
                "where cust.SAKTIF = 1 And dh.Status = 3 " & _
                "and not(di.ItemId='G0003') " & _
                "and cust.KD_PAJAK='{0}' " & _
                "and dh.InvoiceNo='{1}'", tpelanggan.EditValue, dv3(Me.BindingContext(dv3).Position)("nobukti").ToString)

                Dim cmd As OleDbCommand = New OleDbCommand(sql, cn)
                Dim drd As OleDbDataReader = cmd.ExecuteReader

                If drd.HasRows Then
                    If drd.Read Then

                        If Not drd(0).ToString.Equals("") Then
                            nama = drd(2).ToString
                            alamat = drd(3).ToString
                            npwp = drd(4).ToString
                            nama_mobiz = drd(5).ToString
                        End If

                    End If
                End If
                drd.Close()

                If nama.Trim.Length > 0 Then

                    dv3(Me.BindingContext(dv3).Position)("tanggal") = convert_date_to_eng(dv3(Me.BindingContext(dv3).Position)("tanggal"))
                    dv3(Me.BindingContext(dv3).Position)("nama_pkp") = nama
                    dv3(Me.BindingContext(dv3).Position)("alamat_pkp") = alamat
                    dv3(Me.BindingContext(dv3).Position)("npwp_pkp") = npwp
                    dv3(Me.BindingContext(dv3).Position)("nama_mobiz") = nama_mobiz

                    tseri.EditValue = dv3.Count

                Else
                    dv3(Me.BindingContext(dv3).Position).Delete()
                    MsgBox("Data tidak ditemukan")
                End If

            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
            Finally

                If Not cn Is Nothing Then
                    If cn.State = ConnectionState.Open Then
                        cn.Close()
                    End If
                End If
            End Try

        End If
    End Sub

    Private Sub GridView4_RowUpdated(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles GridView4.RowUpdated

        Dim no1 As Integer = dv4(Me.BindingContext(dv4).Position)("noawal").ToString.Length
        Dim no2 As Integer = dv4(Me.BindingContext(dv4).Position)("noakhir").ToString.Length

        If Not (no1 = 0 Or no2 = 0) Then

            Dim noawal As String = dv4(Me.BindingContext(dv4).Position)("noawal").ToString
            noawal = noawal.Substring(noawal.Length - 8)

            Dim noakhir As String = dv4(Me.BindingContext(dv4).Position)("noakhir").ToString
            noakhir = noakhir.Substring(noakhir.Length - 8)

            Dim range As Integer = noakhir - noawal

            dv4(Me.BindingContext(dv4).Position)("norange") = range

        Else
            MsgBox("Lengkapi dulu noseri awal dan noseri akhir")
            dv4(Me.BindingContext(dv4).Position).Delete()
        End If

    End Sub

    Private Sub grid4_KeyDown(sender As Object, e As KeyEventArgs) Handles grid4.KeyDown

        If IsNothing(dv4) Then
            Return
        End If

        If dv4.Count <= 0 Then
            Return
        End If

        If e.KeyCode = Keys.Delete Then

            If MsgBox("Akan dihapus ???", vbQuestion + vbYesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                dv4(Me.BindingContext(dv4).Position).Delete()
            End If

        End If
    End Sub

    Private Sub grid3_KeyDown(sender As Object, e As KeyEventArgs) Handles grid3.KeyDown

        If IsNothing(dv3) Then
            Return
        End If

        If dv3.Count <= 0 Then
            Return
        End If

        If e.KeyCode = Keys.Delete Then

            If MsgBox("Akan dihapus ???", vbQuestion + vbYesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                dv3(Me.BindingContext(dv3).Position).Delete()

                tseri.EditValue = dv3.Count

            End If

        End If
    End Sub

    Private Sub ttgl2_EditValueChanged(sender As Object, e As EventArgs) Handles ttgl2.EditValueChanged

        tmasa.EditValue = Month(ttgl2.EditValue)
        tthn.EditValue = Year(ttgl2.EditValue)

    End Sub

    Private Sub XtraTabControl1_Click(sender As Object, e As EventArgs) Handles XtraTabControl1.Click
        r_cust_pkp()
        load_noseri()
        ttgl1.Focus()
    End Sub
    Private Sub btload_Click(sender As Object, e As EventArgs) Handles btload.Click
        load_data()
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        export_data()
    End Sub


    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click

        Cursor = Cursors.WaitCursor

        Dim cn As OleDbConnection = Nothing

        Try

            cn = New OleDbConnection
            cn = ClassMy.open_conn()

            Dim sql As String = String.Format("select YEAR(tglfak) as thn,MONTH(tglfak) as bln,ItemId,sum(convert(numeric,jml)) as jml,sum(total) as total from v_penjualan where year(tglfak)={0} group by YEAR(tglfak),MONTH(tglfak),ItemId", TextEdit1.EditValue)

            Dim ds As New DataSet
            ds = ClassMy.GetDataSet(sql, cn)

            Dim ds1 As New dt_penjualan
            ds1.Clear()
            ds1.Tables(0).Merge(ds.Tables(0))

            crtotal_perbulan = New Rpenjualan
            crtotal_perbulan.SetDataSource(ds1)

            CrystalReportViewer1.ReportSource = crtotal_perbulan
            CrystalReportViewer1.Refresh()

            Cursor = Cursors.Default

        Catch ex As Exception
            Cursor = Cursors.Default
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            Cursor = Cursors.Default

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If
        End Try

    End Sub


    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click

        open_wait()
        gridtot_detail.DataSource = Nothing

        Dim sql As String = String.Format("select a.InvoiceNo,a.tglfak,a.CustomerId,a.Name,b.Name as nama_brg,convert(numeric,a.jml) as jml,a.total " & _
        "from v_penjualan a inner join v_items b on a.ItemId=b.ItemId where not(a.ItemId='G0003') and a.tglfak>='{0}' and a.tglfak<='{1}'", convert_date_to_eng(ttgl1_cr.EditValue), convert_date_to_eng(ttgl2_cr.EditValue))

        If Not (tnama_cr.EditValue = "") Then
            sql = String.Format(" {0} and a.Name like '%{1}%'", sql, tnama_cr.EditValue)
        End If

        If Not (tbarang_cr.EditValue = "All") Then

            If tbarang_cr.EditValue = "150 ML" Then
                sql = String.Format(" {0} and a.ItemId='{1}'", sql, "C0007")
            End If

            If tbarang_cr.EditValue = "240 ML" Then
                sql = String.Format(" {0} and a.ItemId='{1}'", sql, "C0002")
            End If

            If tbarang_cr.EditValue = "330 ML" Then
                sql = String.Format(" {0} and a.ItemId='{1}'", sql, "B0005")
            End If

            If tbarang_cr.EditValue = "500 ML" Then
                sql = String.Format(" {0} and a.ItemId='{1}'", sql, "B0001")
            End If

            If tbarang_cr.EditValue = "600 ML" Then
                sql = String.Format(" {0} and a.ItemId in ('{1}','{2}','{3}')", sql, "B0003", "B0003A", "B0003N")
            End If

            If tbarang_cr.EditValue = "1500 ML" Then
                sql = String.Format(" {0} and a.ItemId='{1}'", sql, "B0002")
            End If

            If tbarang_cr.EditValue = "19 LTR" Then
                sql = String.Format(" {0} and a.ItemId='{1}'", sql, "G0001")
            End If

        End If

        Dim cn As OleDbConnection = Nothing
        Dim ds As DataSet

        Try

            dv_det = Nothing

            cn = New OleDbConnection
            cn = ClassMy.open_conn

            ds = New DataSet()
            ds = ClassMy.GetDataSet(sql, cn)

            dvmanager_det = New DataViewManager(ds)
            dv_det = dvmanager_det.CreateDataView(ds.Tables(0))

            gridtot_detail.DataSource = dv_det

        Catch ex As OleDb.OleDbException
            close_wait()
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            close_wait()

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If

        End Try

    End Sub


    Private Sub DetailPenjualanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DetailPenjualanToolStripMenuItem.Click
        open_wait()

        XtraTabControl2.SelectedTabPageIndex = 1

        close_wait()
    End Sub

    Private Sub TotalPerbulanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TotalPerbulanToolStripMenuItem.Click

        open_wait()

        XtraTabControl2.SelectedTabPageIndex = 0

        close_wait()

    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click

        open_wait()
        grid_ret.DataSource = Nothing

        Dim sql As String = String.Format("SELECT      sr_a.SalesReturnNo, convert(date,sr_a.TransactionDate) as tgl,bc.BusinessPartnerId,bc.Name, " & _
        "itm.Name as itemname,convert(numeric,sr_b.QtyReturn) as jml,convert(numeric,sr_b.Total) as total " & _
        "FROM         M1Company_Live.Sales.SalesReturnHeader sr_a INNER JOIN " & _
                      "M1Company_Live.Sales.SalesReturnDetail sr_b ON sr_a.SalesReturnNo = sr_b.SalesReturnNo INNER JOIN " & _
                      "M1Company_Live.Shared.CustomFields csf ON sr_a.RowId = csf.RowId " & _
                      "inner join Shared.BusinessPartners bc on sr_a.CustomerId=bc.BusinessPartnerId " & _
                      "inner join Inventory.ItemS itm on sr_b.ItemId=itm.ItemId " & _
        "WHERE     sr_a.Status = 1 and not(sr_b.ItemId='G0003') and convert(date,sr_a.TransactionDate)>='{0}' and convert(date,sr_a.TransactionDate)<='{1}'", convert_date_to_eng(ttgl1_ret.EditValue), convert_date_to_eng(ttgl2_ret.EditValue))

        If Not (tnama_ret.EditValue = "") Then
            sql = String.Format(" {0} and bc.Name like '%{1}%'", sql, tnama_ret.EditValue)
        End If

        If Not (tbarang_ret.EditValue = "All") Then

            If tbarang_ret.EditValue = "150 ML" Then
                sql = String.Format(" {0} and itm.ItemId='{1}'", sql, "C0007")
            End If

            If tbarang_ret.EditValue = "240 ML" Then
                sql = String.Format(" {0} and itm.ItemId='{1}'", sql, "C0002")
            End If

            If tbarang_ret.EditValue = "330 ML" Then
                sql = String.Format(" {0} and itm.ItemId='{1}'", sql, "B0005")
            End If

            If tbarang_ret.EditValue = "500 ML" Then
                sql = String.Format(" {0} and itm.ItemId='{1}'", sql, "B0001")
            End If

            If tbarang_ret.EditValue = "600 ML" Then
                sql = String.Format(" {0} and itm.ItemId in ('{1}','{2}','{3}')", sql, "B0003", "B0003A", "B0003N")
            End If

            If tbarang_ret.EditValue = "1500 ML" Then
                sql = String.Format(" {0} and itm.ItemId='{1}'", sql, "B0002")
            End If

            If tbarang_ret.EditValue = "19 LTR" Then
                sql = String.Format(" {0} and itm.ItemId='{1}'", sql, "G0001")
            End If

        End If

        Dim cn As OleDbConnection = Nothing
        Dim ds As DataSet

        Try

            dv_det = Nothing

            cn = New OleDbConnection
            cn = ClassMy.open_conn_mobiz

            ds = New DataSet()
            ds = ClassMy.GetDataSet(sql, cn)

            dvmanager_ret = New DataViewManager(ds)
            dv_ret = dvmanager_ret.CreateDataView(ds.Tables(0))

            grid_ret.DataSource = dv_ret

        Catch ex As OleDb.OleDbException
            close_wait()
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            close_wait()

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If

        End Try

    End Sub

    Private Sub ReturPenjualanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReturPenjualanToolStripMenuItem.Click
        open_wait()

        XtraTabControl2.SelectedTabPageIndex = 2

        close_wait()
    End Sub

    Private Function ShowSaveFileDialog(ByVal title As String, ByVal filter As String) As String
        Dim dlg As New SaveFileDialog()
        Dim name As String = Application.ProductName
        Dim n As Integer = name.LastIndexOf(".") + 1
        If n > 0 Then
            name = name.Substring(n, name.Length - n)
        End If
        dlg.Title = "Export To " & title
        dlg.FileName = name
        dlg.Filter = filter
        If dlg.ShowDialog() = DialogResult.OK Then
            Return dlg.FileName
        End If
        Return ""
    End Function

    Private Sub OpenFile(ByVal fileName As String)
        If XtraMessageBox.Show("Anda ingin membuka file ?", "Export To...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Dim process As New System.Diagnostics.Process()
                process.StartInfo.FileName = fileName
                process.StartInfo.Verb = "Open"
                process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                process.Start()
            Catch
                DevExpress.XtraEditors.XtraMessageBox.Show(Me, "Data tidak ditemukan", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
        '   progressBarControl1.Position = 0
    End Sub

    Private Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick

        If IsNothing(dv_det) Then
            Return
        End If

        If dv_det.Count <= 0 Then
            Return
        End If

        Dim fileName As String = ShowSaveFileDialog("Excel 2007", "Microsoft Excel|*.xlsx")

        If fileName = String.Empty Then
            Return
        End If

        GridView7.ExportToXlsx(fileName)
        OpenFile(fileName)

    End Sub

    Private Sub BarButtonItem2_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick

        If IsNothing(dv_det) Then
            Return
        End If

        If dv_det.Count <= 0 Then
            Return
        End If

        Dim fileName As String = ShowSaveFileDialog("Text Files", "Text Files|*.txt")

        If fileName = String.Empty Then
            Return
        End If

        GridView7.ExportToText(fileName)
        OpenFile(fileName)

    End Sub


    Private Sub BarButtonItem3_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem3.ItemClick

        If IsNothing(dv_ret) Then
            Return
        End If

        If dv_ret.Count <= 0 Then
            Return
        End If

        Dim fileName As String = ShowSaveFileDialog("Excel 2007", "Microsoft Excel|*.xlsx")

        If fileName = String.Empty Then
            Return
        End If

        GridView9.ExportToXlsx(fileName)
        OpenFile(fileName)

    End Sub

    Private Sub BarButtonItem4_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem4.ItemClick

        If IsNothing(dv_ret) Then
            Return
        End If

        If dv_ret.Count <= 0 Then
            Return
        End If

        Dim fileName As String = ShowSaveFileDialog("Text Files", "Text Files|*.txt")

        If fileName = String.Empty Then
            Return
        End If

        GridView9.ExportToText(fileName)
        OpenFile(fileName)

    End Sub

    Private Sub SimpleButton5_Click(sender As Object, e As EventArgs) Handles SimpleButton5.Click

        Cursor = Cursors.WaitCursor

        Dim cn As OleDbConnection = Nothing

        Try

            cn = New OleDbConnection
            cn = ClassMy.open_conn()

            Dim sql As String = String.Format("SELECT    year(in_h2.TransactionDate) as thn,month(in_h2.TransactionDate) as bln, " & _
            "in_d2.ItemId,convert(int,sum(in_d2.Quantity)) as jml " & _
            "FROM         M1Company_Live.Inventory.InventoryTransferHeader in_h2 INNER JOIN " & _
            "M1Company_Live.Inventory.InventoryTransferDetail in_d2 ON in_h2.InventoryTransferNo = in_d2.InventoryTransferNo LEFT JOIN " & _
            "M1Company_Live.Shared.Persons pkr3 ON in_h2.Text2 = pkr3.PersonId  " & _
            "WHERE     in_h2.WarehouseFromId = 'JABUNG' AND in_h2.WarehouseToId = 'SH' AND in_h2.Status = 1 and " & _
            "in_d2.ItemId in  ('C0007', 'B0002', 'C0002', 'B0005', 'G0001', 'B0001', 'B0003', 'B0003N', 'B0003A','B0003E', 'G0003','B002A','C002A','B005A','B003A') " & _
            "and year(in_h2.TransactionDate)={0} " & _
            "GROUP BY year(in_h2.TransactionDate),month(in_h2.TransactionDate), in_d2.ItemId", TextEdit2.EditValue)

            Dim ds As New DataSet
            ds = ClassMy.GetDataSet(sql, cn)

            Dim ds1 As New dt_pembelian
            ds1.Clear()
            ds1.Tables(0).Merge(ds.Tables(0))

            crtotal_pembelian = New Rpembelian
            crtotal_pembelian.SetDataSource(ds1)

            CrystalReportViewer2.ReportSource = crtotal_pembelian
            CrystalReportViewer2.Refresh()

            Cursor = Cursors.Default

        Catch ex As Exception
            Cursor = Cursors.Default
            MsgBox(ex.ToString, MsgBoxStyle.Information, "Informasi")
        Finally

            Cursor = Cursors.Default

            If Not cn Is Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End If
        End Try

    End Sub

    Private Sub DetailBarangDrJabungToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DetailBarangDrJabungToolStripMenuItem.Click

        open_wait()

        XtraTabControl2.SelectedTabPageIndex = 3

        close_wait()

    End Sub


End Class
