Imports System.IO
Imports System.Drawing.Imaging
Imports System.Math
Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim ofd As New OpenFileDialog()

        ofd.Filter = "画像ファイル(*.png;*.jpg)|*.png;*.jpg"
        ofd.Title = "画像ファイル(.png, .jpg）を選択してください"
        ofd.RestoreDirectory = True
        If ofd.ShowDialog() = DialogResult.OK Then





            Dim currentImage As Image = Image.FromFile(ofd.FileName)
            Dim bm As Bitmap = CType(currentImage, Bitmap)

            Dim size_x As Integer = bm.Width
            Dim size_y As Integer = bm.Height

            Dim c1(size_x, size_y) As Color
            Dim a1(size_x, size_y) As Integer
            Dim r1(size_x, size_y) As Integer
            Dim g1(size_x, size_y) As Integer
            Dim b1(size_x, size_y) As Integer

            '画像の縦横のピクセル数が10の倍数になるようにトリミング
            If size_x Mod 10 <> 0 Then
                Do
                    size_x = size_x - 1
                Loop Until size_x Mod 10 = 0
            End If

            If size_y Mod 10 <> 0 Then
                Do
                    size_y = size_y - 1
                Loop Until size_y Mod 10 = 0
            End If


            If size_x Mod 10 = 0 Then
                If size_y Mod 10 = 0 Then
                    For bmload_x = 0 To size_x - 1
                        For bmload_y = 0 To size_y - 1
                            c1(bmload_x, bmload_y) = bm.GetPixel(bmload_x, bmload_y)
                            a1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).A
                            r1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).R
                            g1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).G
                            b1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).B
                        Next
                    Next

                    Dim img As New Bitmap(size_x, size_y)

                    For d_x_f = 0 To size_x - 1
                        For d_y_f = 0 To size_y - 1
                            img.SetPixel(d_x_f, d_y_f, c1(d_x_f, d_y_f))
                        Next
                    Next

                    Dim fbd As New FolderBrowserDialog

                    fbd.Description = "画像を保存するフォルダを選択してください"

                    If fbd.ShowDialog(Me) = DialogResult.OK Then

                        Dim create_time As String = System.DateTime.Now.ToString("yyMMddHHmmss")

                        Using fs As FileStream = File.Create(fbd.SelectedPath & "\" & create_time & "maze.gif")
                            fs.Close()
                        End Using


                        'refer to https://dobon.net/vb/dotnet/graphics/createanimatedgif.html#section2
                        'Line83 to Line 215, Line324 to Line365, Line446 to Line510, Line517 to Line521

                        Dim wFs As New FileStream(fbd.SelectedPath & "\" & create_time & "maze.gif", FileMode.Truncate, FileAccess.Write, FileShare.None)

                        Dim w As New BinaryWriter(wFs)

                        Dim memostre As New MemoryStream()
                        Dim gct As Boolean = False
                        Dim cTSize As Integer = 0

                        img.Save(memostre, ImageFormat.Gif)
                        memostre.Position = 0

                        Dim byte1(5) As Byte

                        memostre.Read(byte1, 0, 6)

                        For header_counter = 0 To 5
                            w.Write(byte1(header_counter))
                        Next


                        Dim lsd(6) As Byte
                        memostre.Read(lsd, 0, 7)

                        If (lsd(4) And &H80) <> 0 Then
                            cTSize = lsd(4) And &H7
                            gct = True
                        Else
                            gct = False
                        End If

                        lsd(4) = CByte(lsd(4) And &H78)

                        For lsd_counter = 0 To 6
                            w.Write(lsd(lsd_counter))
                        Next


                        Dim AE As Byte() = New Byte(18) {}

                        '拡張導入符 (Extension Introducer)
                        AE(0) = &H21
                        'アプリケーション拡張ラベル (Application Extension Label)
                        AE(1) = &HFF
                        'ブロック寸法 (Block Size)
                        AE(2) = &HB
                        'アプリケーション識別名 (Application Identifier)
                        AE(3) = CByte(AscW("N"c))
                        AE(4) = CByte(AscW("E"c))
                        AE(5) = CByte(AscW("T"c))
                        AE(6) = CByte(AscW("S"c))
                        AE(7) = CByte(AscW("C"c))
                        AE(8) = CByte(AscW("A"c))
                        AE(9) = CByte(AscW("P"c))
                        AE(10) = CByte(AscW("E"c))
                        'アプリケーション確証符号 (Application Authentication Code)
                        AE(11) = CByte(AscW("2"c))
                        AE(12) = CByte(AscW("."c))
                        AE(13) = CByte(AscW("0"c))
                        'データ副ブロック寸法 (Data Sub-block Size)
                        AE(14) = &H3
                        '詰め込み欄 [ネットスケープ拡張コード (Netscape Extension Code)]
                        AE(15) = &H1
                        '繰り返し回数 (Loop Count)
                        Dim loopCountBytes As Byte() = BitConverter.GetBytes(0)
                        AE(16) = loopCountBytes(0)
                        AE(17) = loopCountBytes(1)
                        'ブロック終了符 (Block Terminator)
                        AE(18) = &H0

                        For AE_counter = 0 To 18
                            w.Write(AE(AE_counter))
                        Next


                        Dim cT(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                        If gct Then
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        Dim GCE As Byte() = New Byte(7) {}

                        '拡張導入符 (Extension Introducer)
                        GCE(0) = &H21
                        'グラフィック制御ラベル (Graphic Control Label)
                        GCE(1) = &HF9
                        'ブロック寸法 (Block Size, Byte Size)
                        GCE(2) = &H4
                        '詰め込み欄 (Packed Field)
                        '透過色指標を使う時は+1
                        '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                        GCE(3) = &H0
                        '遅延時間 (Delay Time)
                        Dim delayTimeBytes As Byte() = BitConverter.GetBytes(100)
                        GCE(4) = delayTimeBytes(0)
                        GCE(5) = delayTimeBytes(1)
                        '透過色指標 (Transparency Index, Transparent Color Index)
                        GCE(6) = &H0
                        'ブロック終了符 (Block Terminator)
                        GCE(7) = &H0

                        For GCE_counter = 0 To 7
                            w.Write(GCE(GCE_counter))
                        Next

                        If memostre.GetBuffer()(memostre.Position) = &H21 Then
                            memostre.Position += 8
                        End If

                        Dim iDes(9) As Byte
                        memostre.Read(iDes, 0, 10)

                        If Not gct Then
                            If (iDes(9) And &H80) = 0 Then
                                Throw New Exception("Not found color table.")
                            End If
                            cTSize = iDes(9) And 7
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        iDes(9) = CByte(iDes(9) Or &H80 Or cTSize)
                        For iDes_counter = 0 To 9
                            w.Write(iDes(iDes_counter))
                        Next

                        w.Write(cT)

                        Dim ImageData(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                        memostre.Read(ImageData, 0, CInt(memostre.Length - memostre.Position - 1))
                        w.Write(ImageData)



                        memostre.SetLength(0)


                        Dim c2(size_x / 10, size_y / 10) As Color
                        Dim a2(size_x / 10, size_y / 10) As Integer
                        Dim r2(size_x / 10, size_y / 10) As Integer
                        Dim g2(size_x / 10, size_y / 10) As Integer
                        Dim b2(size_x / 10, size_y / 10) As Integer

                        For x1 = 0 To size_x / 10 - 1
                            For y1 = 0 To size_y / 10 - 1
                                For x2 = 0 To 9
                                    For y2 = 0 To 9
                                        a2(x1, y1) = a2(x1, y1) + a1(x1 * 10 + x2, y1 * 10 + y2)
                                        r2(x1, y1) = r2(x1, y1) + r1(x1 * 10 + x2, y1 * 10 + y2)
                                        g2(x1, y1) = g2(x1, y1) + g1(x1 * 10 + x2, y1 * 10 + y2)
                                        b2(x1, y1) = b2(x1, y1) + b1(x1 * 10 + x2, y1 * 10 + y2)
                                    Next
                                Next
                                a2(x1, y1) = a2(x1, y1) / 100
                                r2(x1, y1) = r2(x1, y1) / 100
                                g2(x1, y1) = g2(x1, y1) / 100
                                b2(x1, y1) = b2(x1, y1) / 100
                                c2(x1, y1) = Color.FromArgb(a2(x1, y1), r2(x1, y1), g2(x1, y1), b2(x1, y1))
                            Next
                        Next


                        For x1 = 0 To size_x / 10 - 1
                            For y1 = 0 To size_y / 10 - 1
                                For x2 = 0 To 9
                                    For y2 = 0 To 9
                                        c1(x1 * 10 + x2, y1 * 10 + y2) = c2(x1, y1)
                                    Next
                                Next
                            Next
                        Next


                        For d_x = 0 To size_x - 1
                            For d_y = 0 To size_y - 1
                                img.SetPixel(d_x, d_y, c1(d_x, d_y))
                            Next
                        Next

                        img.Save(memostre, ImageFormat.Gif)
                        memostre.Position = 0

                        memostre.Position += 6 + 7


                        Dim cT2(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                        If gct Then
                            memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        Dim GCE2 As Byte() = New Byte(7) {}

                        '拡張導入符 (Extension Introducer)
                        GCE2(0) = &H21
                        'グラフィック制御ラベル (Graphic Control Label)
                        GCE2(1) = &HF9
                        'ブロック寸法 (Block Size, Byte Size)
                        GCE2(2) = &H4
                        '詰め込み欄 (Packed Field)
                        '透過色指標を使う時は+1
                        '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                        GCE2(3) = &H0
                        '遅延時間 (Delay Time)
                        Dim delayTimeBytes2 As Byte() = BitConverter.GetBytes(100)
                        GCE2(4) = delayTimeBytes2(0)
                        GCE2(5) = delayTimeBytes2(1)
                        '透過色指標 (Transparency Index, Transparent Color Index)
                        GCE2(6) = &H0
                        'ブロック終了符 (Block Terminator)
                        GCE2(7) = &H0

                        For GCE2_counter = 0 To 7
                            w.Write(GCE2(GCE2_counter))
                        Next

                        If memostre.GetBuffer()(memostre.Position) = &H21 Then
                            memostre.Position += 8
                        End If

                        Dim iDes2(9) As Byte
                        memostre.Read(iDes2, 0, 10)

                        If Not gct Then
                            If (iDes2(9) And &H80) = 0 Then
                                Throw New Exception("Not found color table.")
                            End If
                            cTSize = iDes2(9) And 7
                            memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        iDes2(9) = CByte(iDes2(9) Or &H80 Or cTSize)
                        For iDes2_counter = 0 To 9
                            w.Write(iDes2(iDes2_counter))
                        Next

                        w.Write(cT2)

                        Dim ImageData2(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                        memostre.Read(ImageData2, 0, CInt(memostre.Length - memostre.Position - 1))
                        w.Write(ImageData2)



                        memostre.SetLength(0)

                        Dim a3(size_x / 10, size_y / 10) As Integer
                        Dim r3(size_x / 10, size_y / 10) As Integer
                        Dim g3(size_x / 10, size_y / 10) As Integer
                        Dim b3(size_x / 10, size_y / 10) As Integer

                        For t = 1 To 60

                            a3(0, 0) = (a2(0, 0) + a2(0, 1) + a2(1, 0)) / 3
                            r3(0, 0) = (r2(0, 0) + r2(0, 1) + r2(1, 0)) / 3
                            g3(0, 0) = (g2(0, 0) + g2(0, 1) + g2(1, 0)) / 3
                            b3(0, 0) = (b2(0, 0) + b2(0, 1) + b2(1, 0)) / 3

                            a3(size_x / 10 - 1, 0) = (a2(size_x / 10 - 1, 0) + a2(size_x / 10 - 1, 1) + a2(size_x / 10 - 2, 0)) / 3
                            r3(size_x / 10 - 1, 0) = (r2(size_x / 10 - 1, 0) + r2(size_x / 10 - 1, 1) + r2(size_x / 10 - 2, 0)) / 3
                            g3(size_x / 10 - 1, 0) = (g2(size_x / 10 - 1, 0) + g2(size_x / 10 - 1, 1) + g2(size_x / 10 - 2, 0)) / 3
                            b3(size_x / 10 - 1, 0) = (b2(size_x / 10 - 1, 0) + b2(size_x / 10 - 1, 1) + b2(size_x / 10 - 2, 0)) / 3

                            a3(0, size_y / 10 - 1) = (a2(0, size_y / 10 - 1) + a2(1, size_y / 10 - 1) + a2(0, size_y / 10 - 2)) / 3
                            r3(0, size_y / 10 - 1) = (r2(0, size_y / 10 - 1) + r2(1, size_y / 10 - 1) + r2(0, size_y / 10 - 2)) / 3
                            g3(0, size_y / 10 - 1) = (g2(0, size_y / 10 - 1) + g2(1, size_y / 10 - 1) + g2(0, size_y / 10 - 2)) / 3
                            b3(0, size_y / 10 - 1) = (b2(0, size_y / 10 - 1) + b2(1, size_y / 10 - 1) + b2(0, size_y / 10 - 2)) / 3

                            a3(size_x / 10 - 1, size_y / 10 - 1) = (a2(size_x / 10 - 1, size_y / 10 - 1) + a2(size_x / 10 - 1, size_y / 10 - 2) + a2(size_x / 10 - 2, size_y / 10 - 1)) / 3
                            r3(size_x / 10 - 1, size_y / 10 - 1) = (r2(size_x / 10 - 1, size_y / 10 - 1) + r2(size_x / 10 - 1, size_y / 10 - 2) + r2(size_x / 10 - 2, size_y / 10 - 1)) / 3
                            g3(size_x / 10 - 1, size_y / 10 - 1) = (g2(size_x / 10 - 1, size_y / 10 - 1) + g2(size_x / 10 - 1, size_y / 10 - 2) + g2(size_x / 10 - 2, size_y / 10 - 1)) / 3
                            b3(size_x / 10 - 1, size_y / 10 - 1) = (b2(size_x / 10 - 1, size_y / 10 - 1) + b2(size_x / 10 - 1, size_y / 10 - 2) + b2(size_x / 10 - 2, size_y / 10 - 1)) / 3

                            For h1 = 1 To size_x / 10 - 2
                                a3(h1, 0) = (a2(h1 - 1, 0) + a2(h1, 0) + a2(h1, 1) + a2(h1 + 1, 0)) / 4
                                r3(h1, 0) = (r2(h1 - 1, 0) + r2(h1, 0) + r2(h1, 1) + r2(h1 + 1, 0)) / 4
                                g3(h1, 0) = (g2(h1 - 1, 0) + g2(h1, 0) + g2(h1, 1) + g2(h1 + 1, 0)) / 4
                                b3(h1, 0) = (b2(h1 - 1, 0) + b2(h1, 0) + b2(h1, 1) + b2(h1 + 1, 0)) / 4
                            Next

                            For h2 = 1 To size_x / 10 - 2
                                a3(h2, size_y / 10 - 1) = (a2(h2 - 1, size_y / 10 - 1) + a2(h2, size_y / 10 - 1) + a2(h2, size_y / 10 - 2) + a2(h2 + 1, size_y / 10 - 1)) / 4
                                r3(h2, size_y / 10 - 1) = (r2(h2 - 1, size_y / 10 - 1) + r2(h2, size_y / 10 - 1) + r2(h2, size_y / 10 - 2) + r2(h2 + 1, size_y / 10 - 1)) / 4
                                g3(h2, size_y / 10 - 1) = (g2(h2 - 1, size_y / 10 - 1) + g2(h2, size_y / 10 - 1) + g2(h2, size_y / 10 - 2) + g2(h2 + 1, size_y / 10 - 1)) / 4
                                b3(h2, size_y / 10 - 1) = (b2(h2 - 1, size_y / 10 - 1) + b2(h2, size_y / 10 - 1) + b2(h2, size_y / 10 - 2) + b2(h2 + 1, size_y / 10 - 1)) / 4
                            Next

                            For v1 = 1 To size_y / 10 - 2
                                a3(0, v1) = (a2(0, v1 - 1) + a2(0, v1) + a2(1, v1) + a2(0, v1 + 1)) / 4
                                r3(0, v1) = (r2(0, v1 - 1) + r2(0, v1) + r2(1, v1) + r2(0, v1 + 1)) / 4
                                g3(0, v1) = (g2(0, v1 - 1) + g2(0, v1) + g2(1, v1) + g2(0, v1 + 1)) / 4
                                b3(0, v1) = (b2(0, v1 - 1) + b2(0, v1) + b2(1, v1) + b2(0, v1 + 1)) / 4
                            Next

                            For v2 = 1 To size_y / 10 - 2
                                a3(size_x / 10 - 1, v2) = (a2(size_x / 10 - 1, v2 - 1) + a2(size_x / 10 - 1, v2) + a2(size_x / 10 - 2, v2) + a2(size_x / 10 - 1, v2 + 1)) / 4
                                r3(size_x / 10 - 1, v2) = (r2(size_x / 10 - 1, v2 - 1) + r2(size_x / 10 - 1, v2) + r2(size_x / 10 - 2, v2) + r2(size_x / 10 - 1, v2 + 1)) / 4
                                g3(size_x / 10 - 1, v2) = (g2(size_x / 10 - 1, v2 - 1) + g2(size_x / 10 - 1, v2) + g2(size_x / 10 - 2, v2) + g2(size_x / 10 - 1, v2 + 1)) / 4
                                b3(size_x / 10 - 1, v2) = (b2(size_x / 10 - 1, v2 - 1) + b2(size_x / 10 - 1, v2) + b2(size_x / 10 - 2, v2) + b2(size_x / 10 - 1, v2 + 1)) / 4
                            Next

                            For h3 = 1 To size_x / 10 - 2
                                For v3 = 1 To size_y / 10 - 2
                                    a3(h3, v3) = (a2(h3 - 1, v3) + a2(h3 + 1, v3) + a2(h3, v3 - 1) + a2(h3, v3 + 1) + a2(h3, v3)) / 5
                                    r3(h3, v3) = (r2(h3 - 1, v3) + r2(h3 + 1, v3) + r2(h3, v3 - 1) + r2(h3, v3 + 1) + r2(h3, v3)) / 5
                                    g3(h3, v3) = (g2(h3 - 1, v3) + g2(h3 + 1, v3) + g2(h3, v3 - 1) + g2(h3, v3 + 1) + g2(h3, v3)) / 5
                                    b3(h3, v3) = (b2(h3 - 1, v3) + b2(h3 + 1, v3) + b2(h3, v3 - 1) + b2(h3, v3 + 1) + b2(h3, v3)) / 5
                                Next
                            Next

                            For cast_x = 0 To size_x / 10 - 1
                                For cast_y = 0 To size_y / 10 - 1
                                    a2(cast_x, cast_y) = a3(cast_x, cast_y)
                                    r2(cast_x, cast_y) = r3(cast_x, cast_y)
                                    g2(cast_x, cast_y) = g3(cast_x, cast_y)
                                    b2(cast_x, cast_y) = b3(cast_x, cast_y)
                                    c2(cast_x, cast_y) = Color.FromArgb(a2(cast_x, cast_y), r2(cast_x, cast_y), g2(cast_x, cast_y), b2(cast_x, cast_y))
                                Next
                            Next

                            For x1 = 0 To size_x / 10 - 1
                                For y1 = 0 To size_y / 10 - 1
                                    For x2 = 0 To 9
                                        For y2 = 0 To 9
                                            c1(x1 * 10 + x2, y1 * 10 + y2) = c2(x1, y1)
                                        Next
                                    Next
                                Next
                            Next


                            'gifアニメにする画像のステップ数
                            Dim printcheck As Boolean

                            printcheck = False

                            If t = 1 Then
                                printcheck = True
                            End If

                            If t = 7 Then
                                printcheck = True
                            End If

                            If t = 14 Then
                                printcheck = True
                            End If

                            If t = 28 Then
                                printcheck = True
                            End If

                            If t = 60 Then
                                printcheck = True
                            End If



                            If printcheck = True Then

                                For d_x = 0 To size_x - 1
                                    For d_y = 0 To size_y - 1
                                        img.SetPixel(d_x, d_y, c1(d_x, d_y))
                                    Next
                                Next

                                img.Save(memostre, ImageFormat.Gif)
                                memostre.Position = 0

                                memostre.Position += 6 + 7


                                Dim cT3(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                                If gct Then
                                    memostre.Read(cT3, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                Dim GCE3 As Byte() = New Byte(7) {}

                                '拡張導入符 (Extension Introducer)
                                GCE3(0) = &H21
                                'グラフィック制御ラベル (Graphic Control Label)
                                GCE3(1) = &HF9
                                'ブロック寸法 (Block Size, Byte Size)
                                GCE3(2) = &H4
                                '詰め込み欄 (Packed Field)
                                '透過色指標を使う時は+1
                                '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                                GCE3(3) = &H0
                                '遅延時間 (Delay Time)
                                Dim delayTimeBytes3 As Byte() = BitConverter.GetBytes(100)
                                GCE3(4) = delayTimeBytes3(0)
                                GCE3(5) = delayTimeBytes3(1)
                                '透過色指標 (Transparency Index, Transparent Color Index)
                                GCE3(6) = &H0
                                'ブロック終了符 (Block Terminator)
                                GCE3(7) = &H0

                                For GCE3_counter = 0 To 7
                                    w.Write(GCE3(GCE3_counter))
                                Next

                                If memostre.GetBuffer()(memostre.Position) = &H21 Then
                                    memostre.Position += 8
                                End If

                                Dim iDes3(9) As Byte
                                memostre.Read(iDes3, 0, 10)

                                If Not gct Then
                                    If (iDes3(9) And &H80) = 0 Then
                                        Throw New Exception("Not found color table.")
                                    End If
                                    cTSize = iDes3(9) And 7
                                    memostre.Read(cT3, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                iDes3(9) = CByte(iDes3(9) Or &H80 Or cTSize)
                                For iDes3_counter = 0 To 9
                                    w.Write(iDes3(iDes3_counter))
                                Next

                                w.Write(cT3)

                                Dim ImageData3(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                                memostre.Read(ImageData3, 0, CInt(memostre.Length - memostre.Position - 1))
                                w.Write(ImageData3)



                                memostre.SetLength(0)

                            End If

                        Next


                        w.Write(CByte(&H3B))

                        memostre.Close()
                        w.Close()
                        wFs.Close()
                        System.Media.SystemSounds.Asterisk.Play()
                        MessageBox.Show("画像の作成に成功しました", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                End If
            End If
        End If

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ofd As New OpenFileDialog()

        ofd.Filter = "画像ファイル(*.png;*.jpg)|*.png;*.jpg"
        ofd.Title = "画像ファイル(.png, .jpg）を選択してください"
        ofd.RestoreDirectory = True
        If ofd.ShowDialog() = DialogResult.OK Then





            Dim currentImage As Image = Image.FromFile(ofd.FileName)
            Dim bm As Bitmap = CType(currentImage, Bitmap)

            Dim size_x As Integer = bm.Width
            Dim size_y As Integer = bm.Height

            Dim c1(size_x, size_y) As Color
            Dim a1(size_x, size_y) As Integer
            Dim r1(size_x, size_y) As Integer
            Dim g1(size_x, size_y) As Integer
            Dim b1(size_x, size_y) As Integer


            '画像の縦横のピクセル数が10の倍数になるようにトリミング
            If size_x Mod 10 <> 0 Then
                Do
                    size_x = size_x - 1
                Loop Until size_x Mod 10 = 0
            End If

            If size_y Mod 10 <> 0 Then
                Do
                    size_y = size_y - 1
                Loop Until size_y Mod 10 = 0
            End If


            If size_x Mod 10 = 0 Then
                If size_y Mod 10 = 0 Then
                    For bmload_x = 0 To size_x - 1
                        For bmload_y = 0 To size_y - 1
                            c1(bmload_x, bmload_y) = bm.GetPixel(bmload_x, bmload_y)
                            a1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).A
                            r1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).R
                            g1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).G
                            b1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).B
                        Next
                    Next

                    Dim img As Bitmap
                    img = New Bitmap(size_x, size_y)

                    For d_x_f = 0 To size_x - 1
                        For d_y_f = 0 To size_y - 1
                            img.SetPixel(d_x_f, d_y_f, c1(d_x_f, d_y_f))
                        Next
                    Next

                    Dim fbd As New FolderBrowserDialog

                    fbd.Description = "画像を保存するフォルダを選択してください"

                    If fbd.ShowDialog(Me) = DialogResult.OK Then

                        Dim create_time As String = System.DateTime.Now.ToString("yyMMddHHmmss")

                        Using fs As FileStream = File.Create(fbd.SelectedPath & "\" & create_time & "bara.gif")
                            fs.Close()
                        End Using


                        'refer to https://dobon.net/vb/dotnet/graphics/createanimatedgif.html#section2
                        'Line609 to Line 741, Line934 to Line998, Line934 to Line998, Line1007 to Line1011

                        Dim wFs As New FileStream(fbd.SelectedPath & "\" & create_time & "bara.gif", FileMode.Truncate, FileAccess.Write, FileShare.None)

                        Dim w As New BinaryWriter(wFs)

                        Dim memostre As New MemoryStream()
                        Dim gct As Boolean = False
                        Dim cTSize As Integer = 0

                        img.Save(memostre, ImageFormat.Gif)
                        memostre.Position = 0

                        Dim byte1(5) As Byte

                        memostre.Read(byte1, 0, 6)

                        For header_counter = 0 To 5
                            w.Write(byte1(header_counter))
                        Next


                        Dim lsd(6) As Byte
                        memostre.Read(lsd, 0, 7)

                        If (lsd(4) And &H80) <> 0 Then
                            cTSize = lsd(4) And &H7
                            gct = True
                        Else
                            gct = False
                        End If

                        lsd(4) = CByte(lsd(4) And &H78)

                        For lsd_counter = 0 To 6
                            w.Write(lsd(lsd_counter))
                        Next


                        Dim AE As Byte() = New Byte(18) {}

                        '拡張導入符 (Extension Introducer)
                        AE(0) = &H21
                        'アプリケーション拡張ラベル (Application Extension Label)
                        AE(1) = &HFF
                        'ブロック寸法 (Block Size)
                        AE(2) = &HB
                        'アプリケーション識別名 (Application Identifier)
                        AE(3) = CByte(AscW("N"c))
                        AE(4) = CByte(AscW("E"c))
                        AE(5) = CByte(AscW("T"c))
                        AE(6) = CByte(AscW("S"c))
                        AE(7) = CByte(AscW("C"c))
                        AE(8) = CByte(AscW("A"c))
                        AE(9) = CByte(AscW("P"c))
                        AE(10) = CByte(AscW("E"c))
                        'アプリケーション確証符号 (Application Authentication Code)
                        AE(11) = CByte(AscW("2"c))
                        AE(12) = CByte(AscW("."c))
                        AE(13) = CByte(AscW("0"c))
                        'データ副ブロック寸法 (Data Sub-block Size)
                        AE(14) = &H3
                        '詰め込み欄 [ネットスケープ拡張コード (Netscape Extension Code)]
                        AE(15) = &H1
                        '繰り返し回数 (Loop Count)
                        Dim loopCountBytes As Byte() = BitConverter.GetBytes(0)
                        AE(16) = loopCountBytes(0)
                        AE(17) = loopCountBytes(1)
                        'ブロック終了符 (Block Terminator)
                        AE(18) = &H0

                        For AE_counter = 0 To 18
                            w.Write(AE(AE_counter))
                        Next


                        Dim cT(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                        If gct Then
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        Dim GCE As Byte() = New Byte(7) {}

                        '拡張導入符 (Extension Introducer)
                        GCE(0) = &H21
                        'グラフィック制御ラベル (Graphic Control Label)
                        GCE(1) = &HF9
                        'ブロック寸法 (Block Size, Byte Size)
                        GCE(2) = &H4
                        '詰め込み欄 (Packed Field)
                        '透過色指標を使う時は+1
                        '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                        GCE(3) = &H0
                        '遅延時間 (Delay Time)
                        Dim delayTimeBytes As Byte() = BitConverter.GetBytes(100)
                        GCE(4) = delayTimeBytes(0)
                        GCE(5) = delayTimeBytes(1)
                        '透過色指標 (Transparency Index, Transparent Color Index)
                        GCE(6) = &H0
                        'ブロック終了符 (Block Terminator)
                        GCE(7) = &H0

                        For GCE_counter = 0 To 7
                            w.Write(GCE(GCE_counter))
                        Next

                        If memostre.GetBuffer()(memostre.Position) = &H21 Then
                            memostre.Position += 8
                        End If

                        Dim iDes(9) As Byte
                        memostre.Read(iDes, 0, 10)

                        If Not gct Then
                            If (iDes(9) And &H80) = 0 Then
                                Throw New Exception("Not found color table.")
                            End If
                            cTSize = iDes(9) And 7
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        iDes(9) = CByte(iDes(9) Or &H80 Or cTSize)
                        For iDes_counter = 0 To 9
                            w.Write(iDes(iDes_counter))
                        Next

                        w.Write(cT)

                        Dim ImageData(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                        memostre.Read(ImageData, 0, CInt(memostre.Length - memostre.Position - 1))
                        w.Write(ImageData)



                        memostre.SetLength(0)





                        Dim c2(size_x, size_y) As Color


                        'tは計算ステップ
                        For t = 0 To 320

                            Dim seed As New System.Random(t + 1)

                            Dim rand As Integer

                            rand = seed.Next(2)

                            If rand = 0 Then
                                For x1 = 0 To 9
                                    For y1 = 0 To 9
                                        c2(x1, y1) = c1(x1 + 10, y1)
                                        c1(x1 + 10, y1) = c1(x1, y1)
                                        c1(x1, y1) = c2(x1, y1)
                                    Next
                                Next
                            Else
                                For x1 = 0 To 9
                                    For y1 = 0 To 9
                                        c2(x1, y1) = c1(x1, y1 + 10)
                                        c1(x1, y1 + 10) = c1(x1, y1)
                                        c1(x1, y1) = c2(x1, y1)
                                    Next
                                Next
                            End If



                            rand = seed.Next(2)

                            If rand = 0 Then
                                For x2 = 0 To 9
                                    For y2 = 0 To 9
                                        c2(x2, y2 + size_y - 10) = c1(x2 + 10, y2 + size_y - 10)
                                        c1(x2 + 10, y2 + size_y - 10) = c1(x2, y2 + size_y - 10)
                                        c1(x2, y2 + size_y - 10) = c2(x2, y2 + size_y - 10)
                                    Next
                                Next
                            Else
                                For x2 = 0 To 9
                                    For y2 = 0 To 9
                                        c2(x2, y2 + size_y - 10) = c1(x2, y2 + size_y - 20)
                                        c1(x2, y2 + size_y - 20) = c1(x2, y2 + size_y - 10)
                                        c1(x2, y2 + size_y - 10) = c2(x2, y2 + size_y - 10)
                                    Next
                                Next
                            End If


                            rand = seed.Next(2)

                            If rand = 0 Then
                                For x3 = 0 To 9
                                    For y3 = 0 To 9
                                        c2(x3 + size_x - 10, y3) = c1(x3 + size_x - 20, y3)
                                        c1(x3 + size_x - 20, y3) = c1(x3 + size_x - 10, y3)
                                        c1(x3 + size_x - 10, y3) = c2(x3 + size_x - 10, y3)
                                    Next
                                Next
                            Else
                                For x3 = 0 To 9
                                    For y3 = 0 To 9
                                        c2(x3 + size_x - 10, y3) = c1(x3 + size_x - 10, y3 + 10)
                                        c1(x3 + size_x - 10, y3 + 10) = c1(x3 + size_x - 10, y3)
                                        c1(x3 + size_x - 10, y3) = c2(x3 + size_x - 10, y3)
                                    Next
                                Next
                            End If


                            rand = seed.Next(2)

                            If rand = 0 Then
                                For x4 = 0 To 9
                                    For y4 = 0 To 9
                                        c2(x4 + size_x - 10, y4 + size_y - 10) = c1(x4 + size_x - 20, y4 + size_y - 10)
                                        c1(x4 + size_x - 20, y4 + size_y - 10) = c1(x4 + size_x - 10, y4 + size_y - 10)
                                        c1(x4 + size_x - 10, y4 + size_y - 10) = c2(x4 + size_x - 10, y4 + size_y - 10)
                                    Next
                                Next
                            Else
                                For x4 = 0 To 9
                                    For y4 = 0 To 9
                                        c2(x4 + size_x - 10, y4 + size_y - 10) = c1(x4 + size_x - 10, y4 + size_y - 20)
                                        c1(x4 + size_x - 10, y4 + size_y - 20) = c1(x4 + size_x - 10, y4 + size_y - 10)
                                        c1(x4 + size_x - 10, y4 + size_y - 10) = c2(x4 + size_x - 10, y4 + size_y - 10)
                                    Next
                                Next
                            End If


                            For change_x = 1 To size_x / 10 - 2
                                For change_y = 1 To size_y / 10 - 2
                                    rand = seed.Next(4)

                                    If rand = 0 Then
                                        For x5 = 0 To 9
                                            For y5 = 0 To 9
                                                c2(change_x * 10 + x5, change_y * 10 + y5) = c1(change_x * 10 + x5 + 10, change_y * 10 + y5)
                                                c1(change_x * 10 + x5 + 10, change_y * 10 + y5) = c1(change_x * 10 + x5, change_y * 10 + y5)
                                                c1(change_x * 10 + x5, change_y * 10 + y5) = c2(change_x * 10 + x5, change_y * 10 + y5)
                                            Next
                                        Next
                                    ElseIf rand = 1 Then
                                        For x5 = 0 To 9
                                            For y5 = 0 To 9
                                                c2(change_x * 10 + x5, change_y * 10 + y5) = c1(change_x * 10 + x5 - 10, change_y * 10 + y5)
                                                c1(change_x * 10 + x5 - 10, change_y * 10 + y5) = c1(change_x * 10 + x5, change_y * 10 + y5)
                                                c1(change_x * 10 + x5, change_y * 10 + y5) = c2(change_x * 10 + x5, change_y * 10 + y5)
                                            Next
                                        Next
                                    ElseIf rand = 2 Then
                                        For x5 = 0 To 9
                                            For y5 = 0 To 9
                                                c2(change_x * 10 + x5, change_y * 10 + y5) = c1(change_x * 10 + x5, change_y * 10 + y5 + 10)
                                                c1(change_x * 10 + x5, change_y * 10 + y5 + 10) = c1(change_x * 10 + x5, change_y * 10 + y5)
                                                c1(change_x * 10 + x5, change_y * 10 + y5) = c2(change_x * 10 + x5, change_y * 10 + y5)
                                            Next
                                        Next
                                    Else
                                        For x5 = 0 To 9
                                            For y5 = 0 To 9
                                                c2(change_x * 10 + x5, change_y * 10 + y5) = c1(change_x * 10 + x5, change_y * 10 + y5 - 10)
                                                c1(change_x * 10 + x5, change_y * 10 + y5 - 10) = c1(change_x * 10 + x5, change_y * 10 + y5)
                                                c1(change_x * 10 + x5, change_y * 10 + y5) = c2(change_x * 10 + x5, change_y * 10 + y5)
                                            Next
                                        Next
                                    End If


                                Next
                            Next






                            'gifアニメにする画像のステップ数
                            Dim printcheck As Boolean

                            printcheck = False

                            If t = 1 Then
                                printcheck = True
                            End If

                            If t = 7 Then
                                printcheck = True
                            End If

                            If t = 28 Then
                                printcheck = True
                            End If

                            If t = 60 Then
                                printcheck = True
                            End If

                            If t = 100 Then
                                printcheck = True
                            End If

                            If t = 160 Then
                                printcheck = True
                            End If

                            If t = 240 Then
                                printcheck = True
                            End If

                            If t = 320 Then
                                printcheck = True
                            End If

                            If printcheck = True Then

                                For d_x = 0 To size_x - 1
                                    For d_y = 0 To size_y - 1
                                        img.SetPixel(d_x, d_y, c1(d_x, d_y))
                                    Next
                                Next

                                img.Save(memostre, ImageFormat.Gif)
                                memostre.Position = 0

                                memostre.Position += 6 + 7


                                Dim cT2(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                                If gct Then
                                    memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                Dim GCE2 As Byte() = New Byte(7) {}

                                '拡張導入符 (Extension Introducer)
                                GCE2(0) = &H21
                                'グラフィック制御ラベル (Graphic Control Label)
                                GCE2(1) = &HF9
                                'ブロック寸法 (Block Size, Byte Size)
                                GCE2(2) = &H4
                                '詰め込み欄 (Packed Field)
                                '透過色指標を使う時は+1
                                '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                                GCE2(3) = &H0
                                '遅延時間 (Delay Time)
                                Dim delayTimeBytes2 As Byte() = BitConverter.GetBytes(100)
                                GCE2(4) = delayTimeBytes2(0)
                                GCE2(5) = delayTimeBytes2(1)
                                '透過色指標 (Transparency Index, Transparent Color Index)
                                GCE2(6) = &H0
                                'ブロック終了符 (Block Terminator)
                                GCE2(7) = &H0

                                For GCE2_counter = 0 To 7
                                    w.Write(GCE2(GCE2_counter))
                                Next

                                If memostre.GetBuffer()(memostre.Position) = &H21 Then
                                    memostre.Position += 8
                                End If

                                Dim iDes2(9) As Byte
                                memostre.Read(iDes2, 0, 10)

                                If Not gct Then
                                    If (iDes2(9) And &H80) = 0 Then
                                        Throw New Exception("Not found color table.")
                                    End If
                                    cTSize = iDes2(9) And 7
                                    memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                iDes2(9) = CByte(iDes2(9) Or &H80 Or cTSize)
                                For iDes2_counter = 0 To 9
                                    w.Write(iDes2(iDes2_counter))
                                Next

                                w.Write(cT2)

                                Dim ImageData2(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                                memostre.Read(ImageData2, 0, CInt(memostre.Length - memostre.Position - 1))
                                w.Write(ImageData2)



                                memostre.SetLength(0)

                            End If




                        Next

                        w.Write(CByte(&H3B))

                        memostre.Close()
                        w.Close()
                        wFs.Close()
                        System.Media.SystemSounds.Asterisk.Play()
                        MessageBox.Show("画像の作成に成功しました", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                End If
            End If



        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ofd As New OpenFileDialog()

        ofd.Filter = "画像ファイル(*.png;*.jpg)|*.png;*.jpg"
        ofd.Title = "画像ファイル(.png, .jpg）を選択してください"
        ofd.RestoreDirectory = True
        If ofd.ShowDialog() = DialogResult.OK Then





            Dim currentImage As Image = Image.FromFile(ofd.FileName)
            Dim bm As Bitmap = CType(currentImage, Bitmap)

            Dim size_x As Integer = bm.Width
            Dim size_y As Integer = bm.Height

            Dim c1(size_x, size_y) As Color
            Dim a1(size_x, size_y) As Integer
            Dim r1(size_x, size_y) As Integer
            Dim g1(size_x, size_y) As Integer
            Dim b1(size_x, size_y) As Integer


            '画像の縦横のピクセル数が10の倍数になるようにトリミング
            If size_x Mod 10 <> 0 Then
                Do
                    size_x = size_x - 1
                Loop Until size_x Mod 10 = 0
            End If

            If size_y Mod 10 <> 0 Then
                Do
                    size_y = size_y - 1
                Loop Until size_y Mod 10 = 0
            End If


            If size_x Mod 10 = 0 Then
                If size_y Mod 10 = 0 Then
                    For bmload_x = 0 To size_x - 1
                        For bmload_y = 0 To size_y - 1
                            c1(bmload_x, bmload_y) = bm.GetPixel(bmload_x, bmload_y)
                            a1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).A
                            r1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).R
                            g1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).G
                            b1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).B
                        Next
                    Next

                    Dim img As Bitmap
                    img = New Bitmap(size_x, size_y)

                    For d_x_f = 0 To size_x - 1
                        For d_y_f = 0 To size_y - 1
                            img.SetPixel(d_x_f, d_y_f, c1(d_x_f, d_y_f))
                        Next
                    Next

                    Dim fbd As New FolderBrowserDialog

                    fbd.Description = "画像を保存するフォルダを選択してください"

                    If fbd.ShowDialog(Me) = DialogResult.OK Then

                        Dim create_time As String = System.DateTime.Now.ToString("yyMMddHHmmss")

                        Using fs As FileStream = File.Create(fbd.SelectedPath & "\" & create_time & "tobi.gif")
                            fs.Close()
                        End Using


                        'refer to https://dobon.net/vb/dotnet/graphics/createanimatedgif.html#section2
                        'Line1099 to Line 1231, Line1312 to Line1376, Line1385 to Line1389

                        Dim wFs As New FileStream(fbd.SelectedPath & "\" & create_time & "tobi.gif", FileMode.Truncate, FileAccess.Write, FileShare.None)

                        Dim w As New BinaryWriter(wFs)

                        Dim memostre As New MemoryStream()
                        Dim gct As Boolean = False
                        Dim cTSize As Integer = 0

                        img.Save(memostre, ImageFormat.Gif)
                        memostre.Position = 0

                        Dim byte1(5) As Byte

                        memostre.Read(byte1, 0, 6)

                        For header_counter = 0 To 5
                            w.Write(byte1(header_counter))
                        Next


                        Dim lsd(6) As Byte
                        memostre.Read(lsd, 0, 7)

                        If (lsd(4) And &H80) <> 0 Then
                            cTSize = lsd(4) And &H7
                            gct = True
                        Else
                            gct = False
                        End If

                        lsd(4) = CByte(lsd(4) And &H78)

                        For lsd_counter = 0 To 6
                            w.Write(lsd(lsd_counter))
                        Next


                        Dim AE As Byte() = New Byte(18) {}

                        '拡張導入符 (Extension Introducer)
                        AE(0) = &H21
                        'アプリケーション拡張ラベル (Application Extension Label)
                        AE(1) = &HFF
                        'ブロック寸法 (Block Size)
                        AE(2) = &HB
                        'アプリケーション識別名 (Application Identifier)
                        AE(3) = CByte(AscW("N"c))
                        AE(4) = CByte(AscW("E"c))
                        AE(5) = CByte(AscW("T"c))
                        AE(6) = CByte(AscW("S"c))
                        AE(7) = CByte(AscW("C"c))
                        AE(8) = CByte(AscW("A"c))
                        AE(9) = CByte(AscW("P"c))
                        AE(10) = CByte(AscW("E"c))
                        'アプリケーション確証符号 (Application Authentication Code)
                        AE(11) = CByte(AscW("2"c))
                        AE(12) = CByte(AscW("."c))
                        AE(13) = CByte(AscW("0"c))
                        'データ副ブロック寸法 (Data Sub-block Size)
                        AE(14) = &H3
                        '詰め込み欄 [ネットスケープ拡張コード (Netscape Extension Code)]
                        AE(15) = &H1
                        '繰り返し回数 (Loop Count)
                        Dim loopCountBytes As Byte() = BitConverter.GetBytes(0)
                        AE(16) = loopCountBytes(0)
                        AE(17) = loopCountBytes(1)
                        'ブロック終了符 (Block Terminator)
                        AE(18) = &H0

                        For AE_counter = 0 To 18
                            w.Write(AE(AE_counter))
                        Next


                        Dim cT(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                        If gct Then
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        Dim GCE As Byte() = New Byte(7) {}

                        '拡張導入符 (Extension Introducer)
                        GCE(0) = &H21
                        'グラフィック制御ラベル (Graphic Control Label)
                        GCE(1) = &HF9
                        'ブロック寸法 (Block Size, Byte Size)
                        GCE(2) = &H4
                        '詰め込み欄 (Packed Field)
                        '透過色指標を使う時は+1
                        '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                        GCE(3) = &H0
                        '遅延時間 (Delay Time)
                        Dim delayTimeBytes As Byte() = BitConverter.GetBytes(100)
                        GCE(4) = delayTimeBytes(0)
                        GCE(5) = delayTimeBytes(1)
                        '透過色指標 (Transparency Index, Transparent Color Index)
                        GCE(6) = &H0
                        'ブロック終了符 (Block Terminator)
                        GCE(7) = &H0

                        For GCE_counter = 0 To 7
                            w.Write(GCE(GCE_counter))
                        Next

                        If memostre.GetBuffer()(memostre.Position) = &H21 Then
                            memostre.Position += 8
                        End If

                        Dim iDes(9) As Byte
                        memostre.Read(iDes, 0, 10)

                        If Not gct Then
                            If (iDes(9) And &H80) = 0 Then
                                Throw New Exception("Not found color table.")
                            End If
                            cTSize = iDes(9) And 7
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        iDes(9) = CByte(iDes(9) Or &H80 Or cTSize)
                        For iDes_counter = 0 To 9
                            w.Write(iDes(iDes_counter))
                        Next

                        w.Write(cT)

                        Dim ImageData(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                        memostre.Read(ImageData, 0, CInt(memostre.Length - memostre.Position - 1))
                        w.Write(ImageData)



                        memostre.SetLength(0)





                        Dim c2(size_x, size_y) As Color


                        'tは計算ステップ
                        For t = 0 To 32000

                            Dim seed As New System.Random(1 + t)

                            Dim rand1 As Integer
                            Dim rand2 As Integer
                            Dim rand3 As Integer
                            Dim rand4 As Integer

                            rand1 = seed.Next(size_x / 10)
                            rand2 = seed.Next(size_y / 10)
                            rand3 = seed.Next(size_x / 10)
                            rand4 = seed.Next(size_y / 10)


                            For x1 = 0 To 9
                                For y1 = 0 To 9
                                    c2(rand1 * 10 + x1, rand2 * 10 + y1) = c1(rand3 * 10 + x1, rand4 * 10 + y1)
                                    c1(rand3 * 10 + x1, rand4 * 10 + y1) = c1(rand1 * 10 + x1, rand2 * 10 + y1)
                                    c1(rand1 * 10 + x1, rand2 * 10 + y1) = c2(rand1 * 10 + x1, rand2 * 10 + y1)
                                Next
                            Next




                            'gifアニメにする画像のステップ数
                            Dim printcheck As Boolean

                            printcheck = False

                            If t = 100 Then
                                printcheck = True
                            End If

                            If t = 700 Then
                                printcheck = True
                            End If

                            If t = 2800 Then
                                printcheck = True
                            End If

                            If t = 6000 Then
                                printcheck = True
                            End If

                            If t = 10000 Then
                                printcheck = True
                            End If

                            If t = 16000 Then
                                printcheck = True
                            End If

                            If t = 24000 Then
                                printcheck = True
                            End If

                            If t = 32000 Then
                                printcheck = True
                            End If

                            If printcheck = True Then

                                For d_x = 0 To size_x - 1
                                    For d_y = 0 To size_y - 1
                                        img.SetPixel(d_x, d_y, c1(d_x, d_y))
                                    Next
                                Next

                                img.Save(memostre, ImageFormat.Gif)
                                memostre.Position = 0

                                memostre.Position += 6 + 7


                                Dim cT2(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                                If gct Then
                                    memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                Dim GCE2 As Byte() = New Byte(7) {}

                                '拡張導入符 (Extension Introducer)
                                GCE2(0) = &H21
                                'グラフィック制御ラベル (Graphic Control Label)
                                GCE2(1) = &HF9
                                'ブロック寸法 (Block Size, Byte Size)
                                GCE2(2) = &H4
                                '詰め込み欄 (Packed Field)
                                '透過色指標を使う時は+1
                                '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                                GCE2(3) = &H0
                                '遅延時間 (Delay Time)
                                Dim delayTimeBytes2 As Byte() = BitConverter.GetBytes(100)
                                GCE2(4) = delayTimeBytes2(0)
                                GCE2(5) = delayTimeBytes2(1)
                                '透過色指標 (Transparency Index, Transparent Color Index)
                                GCE2(6) = &H0
                                'ブロック終了符 (Block Terminator)
                                GCE2(7) = &H0

                                For GCE2_counter = 0 To 7
                                    w.Write(GCE2(GCE2_counter))
                                Next

                                If memostre.GetBuffer()(memostre.Position) = &H21 Then
                                    memostre.Position += 8
                                End If

                                Dim iDes2(9) As Byte
                                memostre.Read(iDes2, 0, 10)

                                If Not gct Then
                                    If (iDes2(9) And &H80) = 0 Then
                                        Throw New Exception("Not found color table.")
                                    End If
                                    cTSize = iDes2(9) And 7
                                    memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                iDes2(9) = CByte(iDes2(9) Or &H80 Or cTSize)
                                For iDes2_counter = 0 To 9
                                    w.Write(iDes2(iDes2_counter))
                                Next

                                w.Write(cT2)

                                Dim ImageData2(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                                memostre.Read(ImageData2, 0, CInt(memostre.Length - memostre.Position - 1))
                                w.Write(ImageData2)



                                memostre.SetLength(0)

                            End If




                        Next

                        w.Write(CByte(&H3B))

                        memostre.Close()
                        w.Close()
                        wFs.Close()
                        System.Media.SystemSounds.Asterisk.Play()
                        MessageBox.Show("画像の作成に成功しました", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                End If
            End If



        End If
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim ofd As New OpenFileDialog()

        ofd.Filter = "画像ファイル(*.png;*.jpg)|*.png;*.jpg"
        ofd.Title = "画像ファイル(.png, .jpg）を選択してください"
        ofd.RestoreDirectory = True
        If ofd.ShowDialog() = DialogResult.OK Then





            Dim currentImage As Image = Image.FromFile(ofd.FileName)
            Dim bm As Bitmap = CType(currentImage, Bitmap)

            Dim size_x As Integer = bm.Width
            Dim size_y As Integer = bm.Height

            Dim c1(size_x, size_y) As Color
            Dim a1(size_x, size_y) As Integer
            Dim r1(size_x, size_y) As Integer
            Dim g1(size_x, size_y) As Integer
            Dim b1(size_x, size_y) As Integer


            '画像の縦横のピクセル数が10の倍数になるようにトリミング
            If size_x Mod 10 <> 0 Then
                Do
                    size_x = size_x - 1
                Loop Until size_x Mod 10 = 0
            End If

            If size_y Mod 10 <> 0 Then
                Do
                    size_y = size_y - 1
                Loop Until size_y Mod 10 = 0
            End If


            If size_x Mod 10 = 0 Then
                If size_y Mod 10 = 0 Then
                    For bmload_x = 0 To size_x - 1
                        For bmload_y = 0 To size_y - 1
                            c1(bmload_x, bmload_y) = bm.GetPixel(bmload_x, bmload_y)
                            a1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).A
                            r1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).R
                            g1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).G
                            b1(bmload_x, bmload_y) = c1(bmload_x, bmload_y).B
                        Next
                    Next

                    Dim img As Bitmap
                    img = New Bitmap(size_x, size_y)

                    For d_x_f = 0 To size_x - 1
                        For d_y_f = 0 To size_y - 1
                            img.SetPixel(d_x_f, d_y_f, c1(d_x_f, d_y_f))
                        Next
                    Next

                    Dim fbd As New FolderBrowserDialog

                    fbd.Description = "画像を保存するフォルダを選択してください"

                    If fbd.ShowDialog(Me) = DialogResult.OK Then

                        Dim create_time As String = System.DateTime.Now.ToString("yyMMddHHmmss")

                        Using fs As FileStream = File.Create(fbd.SelectedPath & "\" & create_time & "guru.gif")
                            fs.Close()
                        End Using


                        'refer to https://dobon.net/vb/dotnet/graphics/createanimatedgif.html#section2
                        'Line1479 to Line 1611, Line1725 to Line1789, Line1798 to Line1802

                        Dim wFs As New FileStream(fbd.SelectedPath & "\" & create_time & "guru.gif", FileMode.Truncate, FileAccess.Write, FileShare.None)

                        Dim w As New BinaryWriter(wFs)

                        Dim memostre As New MemoryStream()
                        Dim gct As Boolean = False
                        Dim cTSize As Integer = 0

                        img.Save(memostre, ImageFormat.Gif)
                        memostre.Position = 0

                        Dim byte1(5) As Byte

                        memostre.Read(byte1, 0, 6)

                        For header_counter = 0 To 5
                            w.Write(byte1(header_counter))
                        Next


                        Dim lsd(6) As Byte
                        memostre.Read(lsd, 0, 7)

                        If (lsd(4) And &H80) <> 0 Then
                            cTSize = lsd(4) And &H7
                            gct = True
                        Else
                            gct = False
                        End If

                        lsd(4) = CByte(lsd(4) And &H78)

                        For lsd_counter = 0 To 6
                            w.Write(lsd(lsd_counter))
                        Next


                        Dim AE As Byte() = New Byte(18) {}

                        '拡張導入符 (Extension Introducer)
                        AE(0) = &H21
                        'アプリケーション拡張ラベル (Application Extension Label)
                        AE(1) = &HFF
                        'ブロック寸法 (Block Size)
                        AE(2) = &HB
                        'アプリケーション識別名 (Application Identifier)
                        AE(3) = CByte(AscW("N"c))
                        AE(4) = CByte(AscW("E"c))
                        AE(5) = CByte(AscW("T"c))
                        AE(6) = CByte(AscW("S"c))
                        AE(7) = CByte(AscW("C"c))
                        AE(8) = CByte(AscW("A"c))
                        AE(9) = CByte(AscW("P"c))
                        AE(10) = CByte(AscW("E"c))
                        'アプリケーション確証符号 (Application Authentication Code)
                        AE(11) = CByte(AscW("2"c))
                        AE(12) = CByte(AscW("."c))
                        AE(13) = CByte(AscW("0"c))
                        'データ副ブロック寸法 (Data Sub-block Size)
                        AE(14) = &H3
                        '詰め込み欄 [ネットスケープ拡張コード (Netscape Extension Code)]
                        AE(15) = &H1
                        '繰り返し回数 (Loop Count)
                        Dim loopCountBytes As Byte() = BitConverter.GetBytes(0)
                        AE(16) = loopCountBytes(0)
                        AE(17) = loopCountBytes(1)
                        'ブロック終了符 (Block Terminator)
                        AE(18) = &H0

                        For AE_counter = 0 To 18
                            w.Write(AE(AE_counter))
                        Next


                        Dim cT(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                        If gct Then
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        Dim GCE As Byte() = New Byte(7) {}

                        '拡張導入符 (Extension Introducer)
                        GCE(0) = &H21
                        'グラフィック制御ラベル (Graphic Control Label)
                        GCE(1) = &HF9
                        'ブロック寸法 (Block Size, Byte Size)
                        GCE(2) = &H4
                        '詰め込み欄 (Packed Field)
                        '透過色指標を使う時は+1
                        '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                        GCE(3) = &H0
                        '遅延時間 (Delay Time)
                        Dim delayTimeBytes As Byte() = BitConverter.GetBytes(100)
                        GCE(4) = delayTimeBytes(0)
                        GCE(5) = delayTimeBytes(1)
                        '透過色指標 (Transparency Index, Transparent Color Index)
                        GCE(6) = &H0
                        'ブロック終了符 (Block Terminator)
                        GCE(7) = &H0

                        For GCE_counter = 0 To 7
                            w.Write(GCE(GCE_counter))
                        Next

                        If memostre.GetBuffer()(memostre.Position) = &H21 Then
                            memostre.Position += 8
                        End If

                        Dim iDes(9) As Byte
                        memostre.Read(iDes, 0, 10)

                        If Not gct Then
                            If (iDes(9) And &H80) = 0 Then
                                Throw New Exception("Not found color table.")
                            End If
                            cTSize = iDes(9) And 7
                            memostre.Read(cT, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                        End If

                        iDes(9) = CByte(iDes(9) Or &H80 Or cTSize)
                        For iDes_counter = 0 To 9
                            w.Write(iDes(iDes_counter))
                        Next

                        w.Write(cT)

                        Dim ImageData(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                        memostre.Read(ImageData, 0, CInt(memostre.Length - memostre.Position - 1))
                        w.Write(ImageData)



                        memostre.SetLength(0)





                        Dim c2(size_x, size_y) As Color


                        'tは計算ステップ
                        For t = 0 To 2500

                            Dim rotate As Integer

                            If size_x > size_y Then
                                rotate = Floor(size_y / 2)
                            Else
                                rotate = Floor(size_x / 2)
                            End If

                            For rotate_counter = 0 To rotate - 1

                                For up = rotate_counter To size_y - 2 - rotate_counter
                                    c2(rotate_counter, up + 1) = c1(rotate_counter, up)
                                Next

                                For leftt = rotate_counter To size_x - 2 - rotate_counter
                                    c2(leftt + 1, size_y - 1 - rotate_counter) = c1(leftt, size_y - 1 - rotate_counter)
                                Next

                                For down = rotate_counter To size_y - 2 - rotate_counter
                                    c2(size_x - 1 - rotate_counter, size_y - 2 - down) = c1(size_x - 1 - rotate_counter, size_y - 1 - down)
                                Next

                                For rightt = rotate_counter To size_x - 2 - rotate_counter
                                    c2(size_x - 2 - rightt, rotate_counter) = c1(size_x - 1 - rightt, rotate_counter)
                                Next


                                For up = rotate_counter To size_y - 2 - rotate_counter
                                    c1(rotate_counter, up + 1) = c2(rotate_counter, up + 1)
                                Next

                                For leftt = rotate_counter To size_x - 2 - rotate_counter
                                    c1(leftt + 1, size_y - 1 - rotate_counter) = c2(leftt + 1, size_y - 1 - rotate_counter)
                                Next

                                For down = rotate_counter To size_y - 2 - rotate_counter
                                    c1(size_x - 1 - rotate_counter, size_y - 2 - down) = c2(size_x - 1 - rotate_counter, size_y - 2 - down)
                                Next

                                For rightt = rotate_counter To size_x - 2 - rotate_counter
                                    c1(size_x - 2 - rightt, rotate_counter) = c2(size_x - 2 - rightt, rotate_counter)
                                Next

                            Next








                            'gifアニメにする画像のステップ数
                            Dim printcheck As Boolean

                            printcheck = False

                            If t = 30 Then
                                printcheck = True
                            End If

                            If t = 100 Then
                                printcheck = True
                            End If

                            If t = 200 Then
                                printcheck = True
                            End If

                            If t = 350 Then
                                printcheck = True
                            End If

                            If t = 550 Then
                                printcheck = True
                            End If

                            If t = 750 Then
                                printcheck = True
                            End If

                            If t = 1000 Then
                                printcheck = True
                            End If

                            If t = 1500 Then
                                printcheck = True
                            End If

                            If t = 2500 Then
                                printcheck = True
                            End If


                            If printcheck = True Then

                                For d_x = 0 To size_x - 1
                                    For d_y = 0 To size_y - 1
                                        img.SetPixel(d_x, d_y, c1(d_x, d_y))
                                    Next
                                Next

                                img.Save(memostre, ImageFormat.Gif)
                                memostre.Position = 0

                                memostre.Position += 6 + 7


                                Dim cT2(CInt(Math.Pow(2, cTSize + 1)) * 3 - 1) As Byte
                                If gct Then
                                    memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                Dim GCE2 As Byte() = New Byte(7) {}

                                '拡張導入符 (Extension Introducer)
                                GCE2(0) = &H21
                                'グラフィック制御ラベル (Graphic Control Label)
                                GCE2(1) = &HF9
                                'ブロック寸法 (Block Size, Byte Size)
                                GCE2(2) = &H4
                                '詰め込み欄 (Packed Field)
                                '透過色指標を使う時は+1
                                '消去方法:そのまま残す+4、背景色でつぶす+8、直前の画像に戻す+12
                                GCE2(3) = &H0
                                '遅延時間 (Delay Time)
                                Dim delayTimeBytes2 As Byte() = BitConverter.GetBytes(100)
                                GCE2(4) = delayTimeBytes2(0)
                                GCE2(5) = delayTimeBytes2(1)
                                '透過色指標 (Transparency Index, Transparent Color Index)
                                GCE2(6) = &H0
                                'ブロック終了符 (Block Terminator)
                                GCE2(7) = &H0

                                For GCE2_counter = 0 To 7
                                    w.Write(GCE2(GCE2_counter))
                                Next

                                If memostre.GetBuffer()(memostre.Position) = &H21 Then
                                    memostre.Position += 8
                                End If

                                Dim iDes2(9) As Byte
                                memostre.Read(iDes2, 0, 10)

                                If Not gct Then
                                    If (iDes2(9) And &H80) = 0 Then
                                        Throw New Exception("Not found color table.")
                                    End If
                                    cTSize = iDes2(9) And 7
                                    memostre.Read(cT2, 0, CInt(Math.Pow(2, cTSize + 1)) * 3)
                                End If

                                iDes2(9) = CByte(iDes2(9) Or &H80 Or cTSize)
                                For iDes2_counter = 0 To 9
                                    w.Write(iDes2(iDes2_counter))
                                Next

                                w.Write(cT2)

                                Dim ImageData2(CInt(memostre.Length - memostre.Position - 1) - 1) As Byte
                                memostre.Read(ImageData2, 0, CInt(memostre.Length - memostre.Position - 1))
                                w.Write(ImageData2)



                                memostre.SetLength(0)

                            End If




                        Next

                        w.Write(CByte(&H3B))

                        memostre.Close()
                        w.Close()
                        wFs.Close()
                        System.Media.SystemSounds.Asterisk.Play()
                        MessageBox.Show("画像の作成に成功しました", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                End If
            End If



        End If

    End Sub


    Private originalSize As Size

    'from https://blog.goo.ne.jp/ashm314/e/8122e3c0b3338371b44b8d5c9822a2e4
    'Line 1820 to Line 1842
    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) _
               Handles MyBase.SizeChanged
        Me.SuspendLayout()

        If Not (Me.WindowState = FormWindowState.Minimized) Then

            Dim sfWidth As Single = (Me.ClientSize.Width / Me.originalSize.Width)
            Dim sfHeight As Single = (Me.ClientSize.Height / Me.originalSize.Height)
            Dim sizeFactor As New SizeF(sfWidth, sfHeight)

            For Each ctrl As Control In Me.Controls


                Dim fntScale As Single = (ctrl.Font.Size * sizeFactor.Height)
                ctrl.Font _
                          = New Font(ctrl.Font.FontFamily, fntScale, ctrl.Font.Style, ctrl.Font.Unit)

                ctrl.Scale(sizeFactor)
            Next

            Me.originalSize = Me.ClientSize
        End If
    End Sub
End Class