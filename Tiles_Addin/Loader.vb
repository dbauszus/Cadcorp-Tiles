Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.SqlClient

<GisLinkProgram("Tiles Addin")> _
Public Class Loader

    Private Shared APP As SisApplication
    Private Shared _sis As MapModeller

    Public Shared Property SIS As MapModeller
        Get
            If _sis Is Nothing Then _sis = APP.TakeoverMapManager
            Return _sis
        End Get
        Set(ByVal value As MapModeller)
            _sis = value
        End Set
    End Property

    Public Sub New(ByVal SISApplication As SisApplication)
        APP = SISApplication

        Dim group As SisRibbonGroup = APP.RibbonGroup
        group.Text = "TILES"

        Dim btnTileIndex As SisRibbonButton = New SisRibbonButton("Crear Index", New SisClickHandler(AddressOf TileIndex))
        btnTileIndex.LargeImage = True
        btnTileIndex.Icon = My.Resources.INDEX
        btnTileIndex.Help = "Crear Index"
        group.Controls.Add(btnTileIndex)

        Dim btnTileCache As SisRibbonButton = New SisRibbonButton("Guardar Index", New SisClickHandler(AddressOf TileCache))
        btnTileCache.LargeImage = True
        btnTileCache.Icon = My.Resources.STORE_INDEX
        btnTileCache.Help = "Guardar Index"
        group.Controls.Add(btnTileCache)

    End Sub

    Private Class Tile

        Public zoom As Integer
        Public xName As Integer
        Public yName As Integer
        Public lon As Double
        Public lat As Double
        Public x1 As Double
        Public x2 As Double
        Public y1 As Double
        Public y2 As Double

    End Class

    Private currentTile As Tile

    Private Sub TileIndex(ByVal sender As Object, ByVal e As SisClickArgs)

        Try
            SIS = e.MapModeller
            Dim x1, x2, y1, y2, z, lon, lat As Double
            SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetViewExtent())
            SIS.SplitPos(lat, lon, z, SIS.GetLatLonHgtFromAxes(x1, y2, 0, "OGC.WGS_1984"))

            'calculate expected number of tiles
            Dim xSizeTile = 38.2185
            Dim ySizeTile = 38.2185
            Dim xSize = Math.Abs(x1 - x2)
            Dim ySize = Math.Abs(y1 - y2)
            Dim expectedTiles As Integer
            For i = 20 To 5 Step -1
                expectedTiles += Math.Ceiling(xSize / xSizeTile) * Math.Ceiling(ySize / ySizeTile)
                xSizeTile *= 2
                ySizeTile *= 2
            Next i

            'set starting zoom level
            Dim zoomStart = 5

            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
            Dim Progress As New Progress()
            Progress.TopMost = True
            Progress.StartPosition = FormStartPosition.CenterScreen
            Progress.ProgressBar.Maximum = expectedTiles
            Progress.ProgressBar.Value = 0
            Progress.Show()

            'calculate tiles
            currentTile = New Tile
            For iScale = zoomStart To 20
                currentTile.zoom = iScale
                currentTile.yName = CLng(Math.Floor((1 - Math.Log(Math.Tan(lat * Math.PI / 180) + 1 / Math.Cos(lat * Math.PI / 180)) / Math.PI) / 2 * 2 ^ currentTile.zoom))
                Do
                    currentTile.xName = CLng(Math.Floor((lon + 180) / 360 * 2 ^ currentTile.zoom))
                    Do
                        CalcTile()
                        SIS.CreateRectangle(currentTile.x1, currentTile.y1, currentTile.x2, currentTile.y2)
                        SIS.SetStr(SIS_OT_CURITEM, 0, "tile$", currentTile.yName.ToString + ".png")
                        SIS.SetStr(SIS_OT_CURITEM, 0, "path$", "\" + currentTile.zoom.ToString + "\" + currentTile.xName.ToString)
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "{""Brush"":{""Style"":""Solid"",""Colour"":{""RGBA"":[255,255,255,255]}}}")
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_layer$", currentTile.zoom.ToString)
                        SIS.UpdateItem()
                        currentTile.xName += 1
                        Try
                            Progress.ProgressBar.Value += 1
                        Catch
                        End Try
                        Application.DoEvents()
                    Loop Until currentTile.x2 > x2
                    currentTile.yName += 1
                Loop Until currentTile.y1 < y1
            Next iScale
            Progress.Dispose()
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            SIS.Dispose()
            SIS = Nothing

        Catch ex As Exception
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CalcTile()
        Dim z = 0
        Dim n = 2 ^ currentTile.zoom
        currentTile.lon = currentTile.xName / n * 360 - 180
        Dim latRad = Math.Atan(Math.Sinh(Math.PI * (1 - 2 * currentTile.yName / n)))
        currentTile.lat = latRad * 180 / Math.PI
        SIS.SplitPos(currentTile.x1, currentTile.y2, z, SIS.GetAxesFromLatLonHgt(currentTile.lat, currentTile.lon, 0, "OGC.WGS_1984"))
        latRad = Math.Atan(Math.Sinh(Math.PI * (1 - 2 * (currentTile.yName + 1) / n)))
        SIS.SplitPos(currentTile.x2, currentTile.y1, z, SIS.GetAxesFromLatLonHgt(latRad * 180 / Math.PI, (currentTile.xName + 1) / n * 360 - 180, 0, "OGC.WGS_1984"))
    End Sub

    Private Sub TileCache(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            SIS.CreateListFromSelection("list")
            Dim folderDlg As New FolderBrowserDialog
            folderDlg.ShowNewFolderButton = True
            Dim Progress As New Progress()
            If (folderDlg.ShowDialog() = DialogResult.OK) Then
                Progress.TopMost = True
                Progress.StartPosition = FormStartPosition.CenterScreen
                Progress.ProgressBar.Maximum = SIS.GetListSize("list")
                Progress.ProgressBar.Value = 0
                Progress.Show()
                Dim selectedPath = folderDlg.SelectedPath
                For i = 0 To SIS.GetListSize("list") - 1
                    SIS.DeselectAll()
                    SIS.OpenList("list", i)
                    SIS.SelectItem()
                    SIS.DoCommand("AComZoomSelect")
                    Directory.CreateDirectory(selectedPath & SIS.GetStr(SIS_OT_CURITEM, 0, "path$"))
                    Dim file = selectedPath & SIS.GetStr(SIS_OT_CURITEM, 0, "path$") & "\" & SIS.GetStr(SIS_OT_CURITEM, 0, "tile$")
                    SIS.ExportRaster("PNG_GDALExporter", file, "width=256,height=256,WORLDFILE=FALSE,ALPHA=TRUE")
                    Progress.ProgressBar.Value += 1
                    Application.DoEvents()
                Next
            End If
            Progress.Dispose()
            SIS.Dispose()
            SIS = Nothing

        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub EmptyList(ByVal List As String)
        Try
            SIS.EmptyList(List)
        Catch
        End Try
    End Sub

End Class