Imports OpenTK
Imports OpenTK.Graphics

Public Class Form1

    '3Dを描画する基本クラス
    Private v3d As View3D

    'サンプルの色付き三角形オブジェクト
    '回転させるために、変数を宣言
    Private boColorTriangle As View3D.BufferObject

    'サンプルのサイコロオブジェクト
    '回転させるために、変数を宣言
    'Private boDice As View3D.BufferObject

    'フレームレートを表示するためのテクスチャ
    'このテクスチャは１秒毎に更新するため、変数を宣言
    Private framerateTexture As Bitmap

    'フレームレートを表示するための四角形オブジェクト
    'ビューポートが変更されたときに位置を再設定する必要があるため、変数を宣言
    Private boFramerate As View3D.BufferObject

    'テクスチャに書き出す時に使うフォント
    Private textureFont As New Font(Me.Font.FontFamily.Name, 30.0F) '30=height 48px


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '私の環境ではOpenTK.GLControlはフォームデザイナでうまく動作しません。
        'ツールボックスに出てこないか、出てきても配置で失敗します。
        'View3DクラスはOpenTK.GLControlから継承しているため、同様にフォームデザイナで使用できません。
        'そのため、ここではForm1上にPanel1を配置し、そこへ以下のコードでView3Dクラスのv3dを配置しています。

        v3d = New View3D()
        AddHandler v3d.LoadModel, AddressOf v3d_LoadModel
        AddHandler v3d.ViewportResized, AddressOf v3d_ViewportResized
        AddHandler v3d.Tick, AddressOf v3d_Tick
        With v3d
            .Dock = System.Windows.Forms.DockStyle.Fill
            '.Location = New System.Drawing.Point(0, 0)
            .Name = "glControl"
            '.Size = New System.Drawing.Size(784, 562)
            .TabIndex = 0
            .VSync = False
            .LimitFPS = 30
        End With
        Panel1.Controls.Add(v3d)
    End Sub


    '3Dビューが初期化されるタイミングで呼び出されます
    '例えばタブビューなどに配置されている場合は、実際に表示されるタイミングまでこの関数は呼び出されません。
    Private Sub v3d_LoadModel(sender As View3D)
        Debug.Print("LoadModel")

        '
        '照明を設定
        '

        '平行光源（手前右上から照らす）
        v3d.SetLight(0, View3D.Light.CreateParallelLight(New Vector3(1.0F, 2.0F, 1.5F), Color4.White, 0.8F))

        '点光源（サイコロの左から照らす）※下記定数のdivを参照のこと
        v3d.SetLight(1, View3D.Light.CreatePointLight(New Vector3(-12.0F, 0.0F, 5.0F), Color4.Red, 0.0F))


        '
        'モデルを追加していく
        '

        'XZ平面と座標軸を書く。単位はcm
        Dim bo As View3D.BufferObject = v3d.Create3DObject(bNormal:=False, bTexture:=False, bColor:=True)
        Dim c As Color4
        For x As Integer = -300 To 300 Step 100
            c = IIf(x = 0, Color4.Red, Color4.Gray)
            bo.AddLine(New Vector3(x, 0, -300), c, New Vector3(x, 0, 300), c) 'Z軸は赤
        Next
        For z As Integer = -300 To 300 Step 100
            c = IIf(z = 0, Color4.Green, Color4.Gray)
            bo.AddLine(New Vector3(-300, 0, z), c, New Vector3(300, 0, z), c) 'X軸は緑
        Next
        c = Color4.Blue
        bo.AddLine(New Vector3(0, -100, 0), c, New Vector3(0, 300, 0), c) 'Y軸は青
        bo.Generate()


        Dim texture As Bitmap
        For i As Integer = 0 To 2 'X,Y,Z

            '座標軸のアルファベット表示用のテクスチャを準備
            texture = New Bitmap(48, 48)
            Dim g As Drawing.Graphics = Drawing.Graphics.FromImage(texture)
            Dim brush As New SolidBrush(Color.FromArgb(0, 0, 0, 0))                         '背景を透明にしたいのでA=0のブラシを作成
            g.CompositingMode = Drawing2D.CompositingMode.SourceCopy                        '透明色を上書き
            g.FillRectangle(brush, 0, 0, texture.Width, texture.Height)
            brush.Dispose()
            g.CompositingMode = Drawing2D.CompositingMode.SourceOver                        '通常モードに戻す
            g.DrawString(Chr(88 + i), textureFont, Brushes.White, New Point(0, 0))              'テキストを描画 X Y Z
            g.Dispose()

            '作成したテクスチャを登録する
            v3d.AddTexture(texture)

            '座標軸のアルファベット表示するためのポイントを準備
            bo = v3d.CreateProintSprite(bTexture:=True, bColor:=False)
            With bo
                .AddPoint(New Vector3(IIf(i = 0, 300, 0), IIf(i = 1, 300, 0), IIf(i = 2, 300, 0)), Color4.Red) 'Colorはダミー
                .Generate()
                .PointSize = 20.0F              'ポイントのサイズを指定
                .PointPerspective = True        '遠近処理をする
                .Lighting = False
                .Blend = True                   'テクスチャを半透明で描く
                .Texture = texture              'テクスチャを関連付け
            End With
        Next

        'FPS表示用のテクスチャを準備
        framerateTexture = New Bitmap(100, 48)   'テクスチャを作成
        v3d.AddTexture(framerateTexture)
        DrawFramerateToTexture()                'テクスチャに描画＆登録

        'フレームレートテクスチャを貼り付けるポリゴンを2Dで作成する
        '2Dで作成するとカメラの位置の影響を受けなくなる
        'この場合の座標系は、+X:右 +Y:上 +Z:手前 になる。
        boFramerate = v3d.Create2DObject(bNormal:=True, bTexture:=True, bColor:=False)
        With boFramerate
            Const w = 25
            Const h = 12
            '画面右上に寄せたいので右上を原点にする
            .AddTriangle(
                New Vector3(-w, 0, 0), New Vector2(0, 0),
                New Vector3(0, 0, 0), New Vector2(1, 0),
                New Vector3(0, -h, 0), New Vector2(1, 1))
            .AddTriangle(
                New Vector3(-w, 0, 0), New Vector2(0, 0),
                New Vector3(0, -h, 0), New Vector2(1, 1),
                New Vector3(-w, -h, 0), New Vector2(0, 1))
            .Generate()
            .Lighting = False
            .Blend = True                   'テクスチャを半透明で描く
            .Texture = framerateTexture     'テクスチャを関連付け
            '画面右上に位置を設定
            SetFrameratePosition()
        End With


        '色付き三角形を描く
        boColorTriangle = v3d.Create3DObject(bNormal:=False, bTexture:=False, bColor:=True)
        With boColorTriangle
            .AddTriangle(
                New Vector3(-10, -5, 0), Color4.Red,
                New Vector3(0, 15, 0), Color4.Yellow,
                New Vector3(10, -5, 0), Color4.Green)
            .Generate()
            .Culling = False '裏面も表示
            .Lighting = False '頂点色を使いたいのでライティングをOFFにする
        End With

        'サイコロを描く
        texture = New Bitmap("..\..\dice.jpg")
        v3d.AddTexture(texture, "dice")
        bo = v3d.Create3DObject(bNormal:=True, bTexture:=True, bColor:=False, "dice") '名前はテクスチャと同じにする必要は無い。管理が違うので同じでも可
        With bo
            Const div = 10 '点光源を平面に当てる場合は、ポリゴンを分割しないとうまく表示されないので、divに1以上を設定して、AddQuad内で分割作成してもらう
            .AddQuad('上面
                New Vector3(-1, 1, -1), New Vector2(0, 0),
                New Vector3(1, 1, -1), New Vector2(0.25, 0),
                New Vector3(1, 1, 1), New Vector2(0.25, 0.33),
                New Vector3(-1, 1, 1), New Vector2(0, 0.33), div)
            .AddQuad('正面
                New Vector3(-1, 1, 1), New Vector2(0, 0.33),
                New Vector3(1, 1, 1), New Vector2(0.25, 0.33),
                New Vector3(1, -1, 1), New Vector2(0.25, 0.66),
                New Vector3(-1, -1, 1), New Vector2(0, 0.66), div)
            .AddQuad('下面
                New Vector3(-1, -1, 1), New Vector2(0, 0.66),
                New Vector3(1, -1, 1), New Vector2(0.25, 0.66),
                New Vector3(1, -1, -1), New Vector2(0.25, 1.0),
                New Vector3(-1, -1, -1), New Vector2(0, 1.0), div)
            .AddQuad('右面
                New Vector3(1, 1, 1), New Vector2(0.25, 0.33),
                New Vector3(1, 1, -1), New Vector2(0.5, 0.33),
                New Vector3(1, -1, -1), New Vector2(0.5, 0.66),
                New Vector3(1, -1, 1), New Vector2(0.25, 0.66), div)
            .AddQuad('裏面
                New Vector3(1, 1, -1), New Vector2(0.5, 0.33),
                New Vector3(-1, 1, -1), New Vector2(0.75, 0.33),
                New Vector3(-1, -1, -1), New Vector2(0.75, 0.66),
                New Vector3(1, -1, -1), New Vector2(0.5, 0.66), div)
            .AddQuad('左面
                New Vector3(-1, 1, -1), New Vector2(0.75, 0.33),
                New Vector3(-1, 1, 1), New Vector2(1.0, 0.33),
                New Vector3(-1, -1, 1), New Vector2(1.0, 0.66),
                New Vector3(-1, -1, -1), New Vector2(0.75, 0.66), div)
            .Generate()
            .Texture = texture
            .Material = View3D.Material.FromPreset(View3D.Material.Preset.PlasticWhite)
            .Material.Ambient = New Color4(0.5F, 0.5F, 0.5F, 1.0F) '上記のプリセットのアンビエントは暗すぎるので、適当に明るくする
            .Lighting = True
            .Martices(0) = Matrix4.CreateScale(10.0F) 'サイコロが小さいので10倍に拡大

            '2個目を表示する。具体的な位置はv3d_Tick()内で設定
            .Martices.Add(Matrix4.Identity)
        End With

    End Sub

    'ビューポートが変更されたときに呼び出されます
    Private Sub v3d_ViewportResized(sender As View3D)
        'フレームレートポリゴンの位置を再設定
        SetFrameratePosition()
    End Sub

    'フレームレートのタイミングで時間的な更新のために呼び出される
    Private Sub v3d_Tick(sender As View3D, elapsedMilliseconds As Long)
        Dim bo As View3D.BufferObject

        '色付き三角形を回転させる
        With boColorTriangle
            Static angTY As Single = 0.0F
            angTY += MathHelper.Pi * elapsedMilliseconds / 1000.0F '1秒で180度回転
            .Martices(0) = Matrix4.CreateRotationY(angTY)
            .Martices(0) *= Matrix4.CreateTranslation(-10.0F, 20.0F, -30.0F) '左上の奥に移動
        End With

        '二つ目のサイコロを回転させる
        bo = v3d.GetBufferObject("dice") 'BufferObjectは変数で保持してもよいし、このように名前で探すこともできる
        If bo IsNot Nothing AndAlso bo.Martices.Count >= 2 Then
            Static angDX As Single = 0.0F
            Static angDY As Single = 0.0F
            angDX += MathHelper.Pi * elapsedMilliseconds / 1000.0F / 6.0F '1秒で30度回転
            angDY += MathHelper.Pi * elapsedMilliseconds / 1000.0F / 9.0F '1秒で20度回転
            bo.Martices(1) = Matrix4.CreateRotationX(angDX)
            bo.Martices(1) *= Matrix4.CreateRotationY(angDY)
            bo.Martices(1) *= Matrix4.CreateScale(5.0F)
            bo.Martices(1) *= Matrix4.CreateTranslation(-30.0F, 0.0F, 0.0F) '5倍拡大、左に移動
        End If

        'フレームレートを表示するテクスチャを更新する
        Static elapsedMillis As Long = 0
        elapsedMillis += elapsedMilliseconds
        If elapsedMillis >= 1000 Then
            '1秒経過した
            elapsedMillis -= 1000
            DrawFramerateToTexture()
        End If
    End Sub


    Private Sub SetFrameratePosition()

        '画面右上に寄せたいので、表示範囲の座標を取得
        Dim rc As RectangleF = v3d.GetVisibleEdgeFor2D(0.0F) 'z=0での表示範囲を得る
        boFramerate.Martices(0) = Matrix4.CreateTranslation(rc.Right, rc.Top, 0.0F) 'ポリゴンを右上に寄せる
    End Sub

    Private Sub DrawFramerateToTexture()

        'テクスチャにフレームレートを書き込む
        Dim g As Drawing.Graphics = Drawing.Graphics.FromImage(framerateTexture)
        Dim brush As New SolidBrush(Color.FromArgb(0, 0, 0, 0))                         '背景を透明にしたいのでA=0のブラシを作成
        g.CompositingMode = Drawing2D.CompositingMode.SourceCopy                        '透明色を上書き
        g.FillRectangle(brush, 0, 0, framerateTexture.Width, framerateTexture.Height)
        brush.Dispose()
        g.CompositingMode = Drawing2D.CompositingMode.SourceOver                        '通常モードに戻す
        g.DrawString(v3d.FPS.ToString("F2"), textureFont, Brushes.Yellow, New Point(0, 0))  'テキストを描画
        g.Dispose()

        '更新したテクスチャを反映させる(登録または更新)
        v3d.UpdateTexture(framerateTexture)
    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click

        '１つのサイコロオブジェクトを保存（配列で渡せるので、複数のオブジェクトを１つのファイルに保存できる）
        Dim bo As View3D.BufferObject = v3d.GetBufferObject("dice")
        If bo IsNot Nothing Then
            Dim fileList As New List(Of String)

            If View3D.BufferObject.SaveToXml({bo}, "Objects.xml", True, "", "Texture", fileList) Then
                Dim s As String = "保存しました"
                For Each path As String In fileList
                    s &= vbNewLine
                    s &= path
                Next
                MsgBox(s)
            Else
                MsgBox("保存できませんでした")
            End If
        Else
            MsgBox("サイコロオブジェクトがありません")
        End If
    End Sub

    Private Sub ButtonLoad_Click(sender As Object, e As EventArgs) Handles ButtonLoad.Click
        Dim arBO As New List(Of View3D.BufferObject)
        Dim arTI As New List(Of View3D.TextureInfo)

        If View3D.BufferObject.LoadFromXml("Objects.xml", arBO, True, "", arTI) Then
            MsgBox(String.Format("{0}個のオブジェクトと{1}個のテクスチャを読み込みました", arBO.Count, arTI.Count))

            For Each bo As View3D.BufferObject In arBO


                Debug.Print("Name of loaded buffer object: " & bo.GetName)

                v3d.AddBufferObject(bo)
                bo.Generate()
            Next

            For Each ti As View3D.TextureInfo In arTI
                v3d.AddTexture(ti.bmp, ti.name)
            Next

        Else
            MsgBox("読み込みに失敗しました")
        End If
    End Sub

    Private Sub ButtonDelete_Click(sender As Object, e As EventArgs) Handles ButtonDelete.Click
        If v3d.GetBufferObject("dice") IsNot Nothing Then
            'サイコロオブジェクトのテクスチャを削除
            v3d.DeleteTexture("dice")

            'サイコロオブジェクトを削除
            v3d.DeleteBufferObject("dice")

            MsgBox("サイコロオブジェクトを削除しました")
        Else
            MsgBox("サイコロオブジェクトはありません")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        v3d.SaveToXml("Env.xml", bView:=True, bCamera:=True, bLight:=True, bOthers:=True)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        v3d.LoadFromXml("Env.xml", bView:=True, bCamera:=True, bLight:=True, bOthers:=True)
    End Sub
End Class
