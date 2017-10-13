Imports TileServer
Imports System.Net
Imports System.Drawing

Public MustInherit Class VirtualEarth

    Implements IServer

    Protected ImageTypeValue As String
    Protected MaximumZoomValue As Integer
    Protected ServerNameValue As String
    Protected URL(3) As String

    Protected Function Tilename(ByVal X As Integer, ByVal Y As Integer, ByVal Zoom As Integer) As String
        Tilename = ""
        Dim N As Integer
        Dim Limit As Integer

        For N = Zoom To 0 Step -1
            Limit = 2 ^ N
            If X < Limit Then
                If Y < Limit Then
                    Tilename = Tilename & "0"
                Else
                    Tilename = Tilename & "2"
                    Y = Y - Limit
                End If
            Else
                If Y < Limit Then
                    Tilename = Tilename & "1"
                Else
                    Tilename = Tilename & "3"
                    Y = Y - Limit
                End If
                X = X - Limit
            End If
        Next

    End Function

    Public Function DownloadTile(ByVal X As Integer, ByVal Y As Integer, ByVal Zoom As Integer, ByVal Filename As String) As Boolean Implements TileServer.IServer.DownloadTile
        Dim XY As Integer = 2 ^ (Zoom + 1)
        X = X Mod XY
        Y = Y Mod XY
        Dim Tile As String = Tilename(X, Y, Zoom)
        Dim S As Integer = (X + Y) Mod 4
        Dim imageURL As String = URL(S) & Tile & ImageTypeValue & "?g=1"

        Dim index As Integer
        Dim img As Image
        Try
            Dim req As Net.HttpWebRequest = DirectCast(Net.HttpWebRequest.Create(imageURL), Net.HttpWebRequest)
            Dim res As Net.HttpWebResponse = DirectCast(req.GetResponse, Net.HttpWebResponse)
            index = res.ContentType.IndexOf("image")
            If index > -1 Then
                img = Image.FromStream(res.GetResponseStream)
                img.Save(Filename)
                DownloadTile = True
            Else
                DownloadTile = False
            End If
            res.Close()
        Catch ex As Exception
            DownloadTile = False
        End Try
    End Function
    Public ReadOnly Property ImageType() As String Implements TileServer.IServer.ImageType
        Get
            ImageType = ImageTypeValue
        End Get
    End Property
    Public ReadOnly Property MaximumZoom() As Integer Implements TileServer.IServer.MaximumZoom
        Get
            MaximumZoom = MaximumZoomValue
        End Get
    End Property
    Public ReadOnly Property ServerName() As String Implements TileServer.IServer.ServerName
        Get
            ServerName = ServerNameValue
        End Get
    End Property
    Protected Property BaseURL0() As String
        Get
            BaseURL0 = URL(0)
        End Get
        Set(ByVal value As String)
            URL(0) = value
        End Set
    End Property
    Protected Property BaseURL1() As String
        Get
            BaseURL1 = URL(1)
        End Get
        Set(ByVal value As String)
            URL(1) = value
        End Set
    End Property
    Protected Property BaseURL2() As String
        Get
            BaseURL2 = URL(2)
        End Get
        Set(ByVal value As String)
            URL(2) = value
        End Set
    End Property
    Protected Property BaseURL3() As String
        Get
            BaseURL3 = URL(3)
        End Get
        Set(ByVal value As String)
            URL(3) = value
        End Set
    End Property
    Protected Property BaseImageType() As String
        Get
            BaseImageType = ImageTypeValue
        End Get
        Set(ByVal value As String)
            ImageTypeValue = value
        End Set
    End Property
    Protected Property BaseServerName() As String
        Get
            BaseServerName = ServerNameValue
        End Get
        Set(ByVal value As String)
            ServerNameValue = value
        End Set
    End Property
    Protected Property BaseMaximumZoom() As Integer
        Get
            BaseMaximumZoom = MaximumZoomValue
        End Get
        Set(ByVal value As Integer)
            MaximumZoomValue = value
        End Set
    End Property
End Class

Public Class VirtualEarthSatellite : Inherits VirtualEarth
    Sub New()
        MyBase.New()
        BaseImageType = ".JPG"
        BaseMaximumZoom = 20
        BaseServerName = "VESatellite"
        BaseURL0 = "http://a0.ortho.tiles.virtualearth.net/tiles/a"
        BaseURL1 = "http://a1.ortho.tiles.virtualearth.net/tiles/a"
        BaseURL2 = "http://a2.ortho.tiles.virtualearth.net/tiles/a"
        BaseURL3 = "http://a3.ortho.tiles.virtualearth.net/tiles/a"
    End Sub
End Class

Public Class VirtualEarthStreetMap : Inherits VirtualEarth
    Sub New()
        MyBase.New()
        BaseImageType = ".PNG"
        BaseMaximumZoom = 18
        BaseServerName = "VEStreetMap"
        BaseURL0 = "http://r0.ortho.tiles.virtualearth.net/tiles/r"
        BaseURL1 = "http://r1.ortho.tiles.virtualearth.net/tiles/r"
        BaseURL2 = "http://r2.ortho.tiles.virtualearth.net/tiles/r"
        BaseURL3 = "http://r3.ortho.tiles.virtualearth.net/tiles/r"
    End Sub
End Class
