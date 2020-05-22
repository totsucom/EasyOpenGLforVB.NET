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
    End Sub

    'フレームレートのタイミングで時間的な更新のために呼び出される
    Private Sub v3d_Tick(sender As View3D, elapsedMilliseconds As Long)
    End Sub
End Class