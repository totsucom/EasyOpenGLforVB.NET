Imports OpenTK
Imports OpenTK.Graphics

Public Class Form1

    '3Dを描画する基本クラス
    Private v3d As View3D

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'View3Dクラスを初期化、フォーム上に配置します
        v3d = New View3D()
        AddHandler v3d.LoadModel, AddressOf v3d_LoadModel
        AddHandler v3d.Tick, AddressOf v3d_Tick
        With v3d
            .Dock = System.Windows.Forms.DockStyle.Fill
            .Name = "glControl"
            .TabIndex = 0
            .VSync = False
            .LimitFPS = 30
        End With
        Panel1.Controls.Add(v3d)
    End Sub

    '3Dビューが初期化されるタイミングで呼び出されます
    Private Sub v3d_LoadModel(sender As View3D)

        '色付き三角形を描く
        Dim bo As View3D.BufferObject = v3d.Create3DObject(bNormal:=False, bTexture:=False, bColor:=True, "Triangle")
        With bo
            .AddTriangle(       '三角形を追加
                New Vector3(-10, -5, 0), Color4.Red,
                New Vector3(0, 15, 0), Color4.Yellow,
                New Vector3(10, -5, 0), Color4.Green)
            .Generate()         '頂点データから頂点バッファ、インデックスバッファを作成
            .Culling = False    '裏面も表示
            .Lighting = False   '頂点色を使いたいのでライティングをOFFにする
        End With

        '座標軸を描く
        bo = v3d.Create3DObject(bNormal:=False, bTexture:=False, bColor:=True, "Axis")
        With bo
            .AddLine(   'X軸
                New Vector3(-100, 0, 0), Color4.Red,
                New Vector3(100, 0, 0), Color4.Red)
            .AddLine(   'Y軸
                New Vector3(0, -100, 0), Color4.Blue,
                New Vector3(0, 100, 0), Color4.Blue)
            .AddLine(   'Z軸
                New Vector3(0, 0, -100), Color4.Green,
                New Vector3(0, 0, 100), Color4.Green)
            .Generate()         '頂点データから頂点バッファ、インデックスバッファを作成
            .Lighting = False   '頂点色を使いたいのでライティングをOFFにする
        End With

        'XZ平面を描く
        bo = v3d.Create3DObject(bNormal:=False, bTexture:=False, bColor:=True, "XZPlane")
        With bo
            .AddQuad(   'XZ平面。頂点に白色、50%の透明度を持たせる
                New Vector3(-50, 0, -50), New Color4(1.0F, 1.0F, 1.0F, 0.5F),
                New Vector3(50, 0, -50), New Color4(1.0F, 1.0F, 1.0F, 0.5F),
                New Vector3(50, 0, 50), New Color4(1.0F, 1.0F, 1.0F, 0.5F),
                New Vector3(-50, 0, 50), New Color4(1.0F, 1.0F, 1.0F, 0.5F))
            .Generate()         '頂点データから頂点バッファ、インデックスバッファを作成
            .Lighting = False   '頂点色を使いたいのでライティングをOFFにする
            .Culling = False    '裏面も表示
            .Blend = True       '半透明処理を行う
        End With
    End Sub

    'フレームレートのタイミングで時間的な更新のために呼び出される
    Private Sub v3d_Tick(sender As View3D, elapsedMilliseconds As Long)

        '名前から登録した三角形を探す
        Dim bo As View3D.BufferObject = v3d.GetBufferObject("Triangle")
        If bo IsNot Nothing Then
            bo.Martices(0) *= Matrix4.CreateRotationY(      'Y軸で回転する行列を作成
                elapsedMilliseconds / 1000.0F * Math.PI)    '1秒間で180度
        End If
    End Sub

End Class