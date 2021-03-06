﻿Imports System.Xml
Imports OpenTK
Imports OpenTK.Graphics
Imports OpenTK.Graphics.OpenGL

Public Class View3D
    Inherits OpenTK.GLControl

    'イベント

    ''' <summary>
    ''' View3D初期化後にユーザー側の初期化処理のために呼ばれます
    ''' これはフォームアプリの起動時ではなく、View3Dが表示されるタイミングで実行されます。
    ''' 例えばタブコントロールの初期状態で非表示のタブに配置した場合、タブが選択されて、表示されるときに呼ばれます。
    ''' 一般にこのイベントで照明、カメラの設定やモデルのロードを行いますが、必ずやらなければならないものはありません。
    ''' 初期化が行われたかどうかを知るには View3D.Loadedプロパティを参照します
    ''' </summary>
    ''' <param name="sender"></param>
    Public Event LoadModel(sender As View3D)

    ''' <summary>
    ''' ウィンドウサイズが変わったときなど、ビューポートが再設定されたときに呼び出されます。
    ''' ビューポートが変化すると、GetVisibleEdgeFor2D()の返す値も変化するため、必要に応じて再計算してください。
    ''' このイベントはフォームアプリの起動時には発生しません。また、View3Dの初期化前にも発生しません。
    ''' </summary>
    ''' <param name="sender"></param>
    Public Event ViewportResized(sender As View3D)

    ''' <summary>
    ''' フレーム描画毎に呼ばれます。一般に時間に応じて行う処理を記述しますが、その内容は特に決まっていません。
    ''' milliSecondsに前回の呼び出しからの経過時間(ミリ秒)が格納されているので、一般にこの値を用いてオブジェクトの移動や回転処理を行います。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="milliSeconds"></param>
    Public Event Tick(sender As View3D, milliSeconds As Long)


    '3D関連
    Private projection As Matrix4                       '視野
    Private cameraMatrix As Matrix4 = Matrix4.Identity  '視点を決める回転行列

    Private mousePos As New Point                       'マウスでカメラを動かすため、座標を記憶
    Private cameraHorizAngle As Single = 0              '視線(カメラの向き)
    Private cameraVertAngle As Single = 0
    Private cameraPosition As New Vector3(0, 0, 0)      '視点(カメラの位置)
    Private cameraDistance As Single = 100              '視点から注目点までの距離

    Private viewingAngleV As Single = MathHelper.PiOver4 '縦方向の視野角
    Private zNear = 1.0F                                '表示する奥行きの範囲。手前
    Private zFar = 10000.0F                             '表示する奥行きの範囲。奥

    Private _clearColor As Color4 = Color4.DarkBlue     'クリアカラー(背景色)


    '表示モデルを管理
    Private arBO As New List(Of BufferObject)

    'テクスチャを管理
    Public Class TextureInfo
        Friend bmp As Bitmap = Nothing
        Friend name As String
        Friend txt As Integer = 0       'バッファID
        Public ReadOnly Property GetBitmap As Bitmap
            Get
                Return bmp
            End Get
        End Property
        Public ReadOnly Property GetName As String
            Get
                Return name
            End Get
        End Property
    End Class
    Private arTxt As New List(Of TextureInfo)

    '照明
    Private _light(3) As Light
    Private _lightNames As Integer() = {LightName.Light0, LightName.Light1, LightName.Light2, LightName.Light3}
    Private _lightCaps As Integer() = {EnableCap.Light0, EnableCap.Light1, EnableCap.Light2, EnableCap.Light3}


    'いろいろ
    Private _loaded As Boolean = False      '初期化後True
    Private _fps As Single                  'フレームレートを保持
    Private _limitFrames As Single = 0      'フレームレート制限値
    Private _waitTimerForFrame As Integer   'フレームレート制限のための待ち時間[ms]

    ''' <summary>
    ''' 初期化されたことを返す
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Loaded As Boolean
        Get
            Return _loaded
        End Get
    End Property

    ''' <summary>
    ''' 現在のフレームレートを取得する
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property FPS As Single
        Get
            Return _fps
        End Get
    End Property

    ''' <summary>
    ''' 最大フレームレートを設定または取得する
    ''' </summary>
    ''' <returns></returns>
    Public Property LimitFPS As Single
        Get
            Return _limitFrames
        End Get
        Set(value As Single)
            If value < 0.1F Then value = 0.1F
            _limitFrames = value
            _waitTimerForFrame = 1000.0F / _limitFrames
        End Set
    End Property

    ''' <summary>
    ''' ３D描画エリアのクリア色、つまり背景色を設定する
    ''' </summary>
    ''' <returns></returns>
    Public Property ClearColor As Color4
        Get
            Return _clearColor
        End Get
        Set(value As Color4)
            If _loaded Then GL.ClearColor(_clearColor)
            _clearColor = value
        End Set
    End Property

    ''' <summary>
    ''' 照明を設定する。indexは照明のインデックスで、０～３を設定。
    ''' 設定を削除したい場合はNothingを渡す
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="value"></param>
    Public Sub SetLight(index As Integer, ByRef value As Light)
        _light(index) = value
        If _loaded AndAlso value IsNot Nothing Then
            'ライトの指定
            GL.Light(_lightNames(index), LightParameter.Position, value.Position)
            GL.Light(_lightNames(index), LightParameter.Ambient, value.Ambient)
            GL.Light(_lightNames(index), LightParameter.Diffuse, value.Diffuse)
            GL.Light(_lightNames(index), LightParameter.Specular, value.Specular)
        End If
    End Sub

    ''' <summary>
    ''' 照明を取得する。取得した照明の個別パラメータを設定しても反映されないので、
    ''' 設定を反映する場合は必ずSetLight()を呼び出すこと
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public Function GetLight(index As Integer) As Light
        Return _light(index)
    End Function

    ''' <summary>
    ''' 実際に使用するにはフォーム上にパネルなどを配置し、Panel1.Controls.Add(View3D変数) などで使用開始する。
    ''' viewingAngleV: 視野の縦角度(ラジアン)を指定。デフォルトは45度。水平方向の視野角はView3Dのアスペクト比で決まる。
    ''' viewDistanceNear: 表示可能な奥行きの範囲。手前側。デフォルトは 1.0F
    ''' viewDistanceFar: 表示可能な奥行きの範囲。奥側。デフォルトは 10000.0F
    ''' useMouse: マウスイベント処理を行い、カメラ移動処理を行うかどうか。デフォルトは行う。
    ''' </summary>
    ''' <param name="viewingAngleV"></param>
    ''' <param name="viewDistanceNear"></param>
    ''' <param name="viewDistanceFar"></param>
    ''' <param name="useMouse"></param>
    Sub New(Optional viewingAngleV As Single = MathHelper.PiOver4,
            Optional viewDistanceNear As Single = 1.0F,
            Optional viewDistanceFar As Single = 10000.0F,
            Optional useMouse As Boolean = True)

        Me.viewingAngleV = viewingAngleV
        If viewDistanceNear < 0.1F Then viewDistanceNear = 0.1F
        If viewDistanceFar < viewDistanceNear Then viewDistanceFar = viewDistanceNear
        zNear = viewDistanceNear
        zFar = viewDistanceFar

        '
        ' イベントハンドラを設定
        '
        AddHandler Me.Load, AddressOf glControl_Load
        AddHandler Me.Resize, AddressOf glControl_Resize
        AddHandler Me.Paint, AddressOf glControl_Paint
        If useMouse Then
            AddHandler Me.MouseDown, AddressOf glControl_MouseDown
            AddHandler Me.MouseMove, AddressOf glControl_MouseMove
            AddHandler Me.MouseWheel, AddressOf glControl_MouseWheel
        End If
        AddHandler Application.Idle, AddressOf Application_Idle
    End Sub

    '3Dビューを継続して更新するために呼び出しています
    Private Sub Application_Idle(sender As Object, e As EventArgs)
        Me.Invalidate() 'このように頻繁に呼ばないとフレームレートを達成できない
    End Sub

    'ビューポートを準備する
    Private Sub PrepareViewport()
        'ビューポートを設定
        GL.Viewport(0, 0, Me.Width, Me.Height)

        '視野
        projection = Matrix4.CreatePerspectiveFieldOfView(
                    viewingAngleV, CSng(Me.Width) / CSng(Me.Height), zNear, zFar)
        GL.MatrixMode(MatrixMode.Projection)
        GL.LoadMatrix(projection)
    End Sub

    Private Sub glControl_Load(ByVal sender As Object, ByVal e As EventArgs)
        '
        ' 3D関連の初期化
        '

        GL.ClearColor(Color4.DarkBlue)
        GL.Enable(EnableCap.DepthTest)

        'ビューポートを準備する
        PrepareViewport()

        'ライトの指定
        For i As Integer = 0 To UBound(_light)
            If _light(i) IsNot Nothing Then
                With _light(i)
                    GL.Light(LightName.Light0, LightParameter.Position, .Position)
                    GL.Light(LightName.Light0, LightParameter.Ambient, .Ambient)
                    GL.Light(LightName.Light0, LightParameter.Diffuse, .Diffuse)
                    GL.Light(LightName.Light0, LightParameter.Specular, .Specular)
                End With
            End If
        Next

        'モデルを生成や初期処理を行うためにユーザー側のハンドラを呼び出す
        _loaded = True
        RaiseEvent LoadModel(Me)
    End Sub

    Private Sub glControl_Resize(ByVal sender As Object, ByVal e As EventArgs)
        If _loaded Then
            'ビューポートを再設定する
            PrepareViewport()
            RaiseEvent ViewportResized(Me)
        End If
    End Sub

    Private Sub glControl_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)

        'FPS制御
        Static swWait As New Stopwatch
        Dim elapsedMilliseconds As Long = 0
        If _limitFrames > 0 AndAlso _waitTimerForFrame > 0 Then
            If Not swWait.IsRunning Then
                swWait.Start()
            Else
                elapsedMilliseconds = swWait.ElapsedMilliseconds
                If elapsedMilliseconds < _waitTimerForFrame Then Exit Sub
                swWait.Restart()
            End If
        End If

        'FPSカウント
        Static count As Integer
        Static swFPS As New Stopwatch
        If Not swFPS.IsRunning Then
            swFPS.Start()
            count = 1
        Else
            count += 1
        End If
        If swFPS.ElapsedMilliseconds >= 1000 Then
            _fps = count / swFPS.ElapsedMilliseconds * 1000.0F
            count = 0
            swFPS.Restart()
        End If



        '画面クリア
        GL.Clear(ClearBufferMask.ColorBufferBit Or ClearBufferMask.DepthBufferBit)

        '視野の設定
        Dim modelview As Matrix4 = Matrix4.LookAt(New Vector3(0, 0, cameraDistance), New Vector3(0, 0, 0), New Vector3(0, 1, 0))
        GL.MatrixMode(MatrixMode.Modelview)
        GL.LoadMatrix(modelview)

        'カリングの共通設定
        GL.CullFace(CullFaceMode.Back)
        GL.FrontFace(FrontFaceDirection.Cw) '時計回りが表

        '照明の共通設定
        For i As Integer = 0 To UBound(_light)
            If _light(i) IsNot Nothing AndAlso _light(i).Active Then
                GL.Enable(_lightCaps(i))
            Else
                GL.Disable(_lightCaps(i))
            End If
        Next
        GL.Enable(EnableCap.Normalize)

        'ブレンディングの共通設定
        GL.BlendFunc(BlendingFactor.SrcAlpha, BlendingFactor.OneMinusSrcAlpha)

        '
        '光源位置の設定
        '

        '視野に追随する光源を設定
        For i As Integer = 0 To UBound(_light)
            If _light(i) IsNot Nothing AndAlso _light(i).FollowCamera Then
                GL.Light(_lightNames(i), LightParameter.Position, _light(i).Position)
            End If
        Next

        GL.PushMatrix()
        GL.MultMatrix(cameraMatrix) 'この行との前後関係が肝

        'カメラ位置の影響を受けない、グローバル位置の光源を設定
        For i As Integer = 0 To UBound(_light)
            If _light(i) IsNot Nothing AndAlso Not _light(i).FollowCamera Then
                GL.Light(_lightNames(i), LightParameter.Position, _light(i).Position)
            End If
        Next
        GL.PopMatrix()


        'デフォルトのマテリアルを記憶
        Dim amb(3) As Single, dif(3) As Single, spe(3) As Single, shi As Single
        GL.GetMaterial(MaterialFace.Front, MaterialParameter.Ambient, amb)
        GL.GetMaterial(MaterialFace.Front, MaterialParameter.Diffuse, dif)
        GL.GetMaterial(MaterialFace.Front, MaterialParameter.Specular, spe)
        GL.GetMaterial(MaterialFace.Front, MaterialParameter.Shininess, shi)

        '描画
        For j As Integer = 0 To 1 '非Blendモデル、次にBlendモデル
            For i As Integer = 0 To arBO.Count - 1
                With arBO(i)
                    If .Active AndAlso .Martices.Count > 0 AndAlso ((j = 0 AndAlso Not .Blend) OrElse (j = 1 AndAlso .Blend)) Then

                        'カリング設定
                        If .Culling Then
                            GL.Enable(EnableCap.CullFace) '裏面は非表示
                        Else
                            GL.Disable(EnableCap.CullFace)
                        End If

                        'ライティング設定
                        If .Lighting Then
                            GL.Enable(EnableCap.Lighting)
                        Else
                            GL.Disable(EnableCap.Lighting)
                        End If

                        'マテリアルを設定
                        If .Material IsNot Nothing Then
                            With .Material
                                GL.Material(MaterialFace.Front, MaterialParameter.Ambient, .Ambient)
                                GL.Material(MaterialFace.Front, MaterialParameter.Diffuse, .Diffuse)
                                GL.Material(MaterialFace.Front, MaterialParameter.Specular, .Specular)
                                GL.Material(MaterialFace.Front, MaterialParameter.Shininess, .Shininess)
                            End With
                        End If

                        'ブレンディング(半透明)を設定
                        If .Blend Then
                            GL.Enable(EnableCap.Blend)
                        Else
                            GL.Disable(EnableCap.Blend)
                        End If

                        'ポイントサイズを設定
                        If ._bm = BeginMode.Points Then GL.PointSize(.PointSize)

                        '頂点バッファを関連付け
                        GL.BindVertexArray(._vao)
                        GL.BindBuffer(BufferTarget.ElementArrayBuffer, ._ibo)

                        'テクスチャを関連付け
                        Dim setTexture As Boolean = False
                        If ._bTexture Then
                            For Each ti As TextureInfo In arTxt
                                If (.Texture IsNot Nothing AndAlso ti.bmp Is .Texture) OrElse
                                   (.TextureName <> "" AndAlso ti.name = .TextureName) Then

                                    GL.BindTexture(TextureTarget.Texture2D, ti.txt)
                                    setTexture = True
                                    Exit For
                                End If
                            Next
                        End If

                        'ポイントスプライトを設定
                        If ._bPointSprite Then
                            GL.Enable(EnableCap.PointSprite)

                            'ポイントにテクスチャ全体を割り当てる
                            If setTexture Then GL.TexEnv(TextureEnvTarget.PointSprite, TextureEnvParameter.CoordReplace, 1.0F) '1.0F=true

                            '遠近感を設定
                            If .PointPerspective Then
                                Static distance As Single() = {}
                                If distance.Count = 0 Then distance = {0, 0, Math.Pow(1 / (projection.M11 * Width / 2), 2)}
                                GL.PointParameter(PointParameterName.PointDistanceAttenuation, distance)
                            Else
                                GL.PointParameter(PointParameterName.PointDistanceAttenuation, New Single() {1, 0, 0})
                            End If
                        End If


                        '各Arrayを有効化
                        GL.EnableClientState(ArrayCap.VertexArray)
                        If ._bNormal Then GL.EnableClientState(ArrayCap.NormalArray)
                        If setTexture Then GL.EnableClientState(ArrayCap.TextureCoordArray)
                        If ._bColor Then GL.EnableClientState(ArrayCap.ColorArray)


                        GL.PushMatrix()
                        If ._bXYZ Then GL.MultMatrix(cameraMatrix) '3Dモデルならカメラ位置を適用

                        'モデルを複数描画
                        For Each m As Matrix4 In .Martices
                            GL.PushMatrix()
                            GL.MultMatrix(m)                    'モデル位置を適用
                            GL.DrawElements(._bm, ._numIndex, DrawElementsType.UnsignedInt, 0) '描画
                            GL.PopMatrix()
                        Next

                        GL.PopMatrix()


                        '各Arrayを無効化
                        GL.DisableClientState(ArrayCap.VertexArray)
                        If ._bNormal Then GL.DisableClientState(ArrayCap.NormalArray)
                        If setTexture Then GL.DisableClientState(ArrayCap.TextureCoordArray)
                        If ._bColor Then GL.DisableClientState(ArrayCap.ColorArray)

                        'バッファの関連付けを解除
                        GL.BindVertexArray(0)
                        GL.BindBuffer(BufferTarget.ElementArrayBuffer, 0)
                        If setTexture Then GL.BindTexture(TextureTarget.Texture2D, 0)

                        'その他、悪さしそうな機能はOFF
                        If ._bPointSprite Then GL.Disable(EnableCap.PointSprite)
                        If .Blend Then GL.Disable(EnableCap.Blend)

                        'マテリアルを戻す
                        If .Material IsNot Nothing Then
                            GL.Material(MaterialFace.Front, MaterialParameter.Ambient, amb)
                            GL.Material(MaterialFace.Front, MaterialParameter.Diffuse, dif)
                            GL.Material(MaterialFace.Front, MaterialParameter.Specular, spe)
                            GL.Material(MaterialFace.Front, MaterialParameter.Shininess, shi)
                        End If
                    End If
                End With
            Next
        Next

        Me.SwapBuffers()

        RaiseEvent Tick(Me, elapsedMilliseconds)
    End Sub

    ''' <summary>
    ''' カメラの位置（視点）を変更する
    ''' position: カメラの移動方向（上下左右ではなく、絶対軸方向）
    ''' relative: Trueで現在座標に加算、Falseで現在座標に置き換える
    ''' </summary>
    ''' <param name="position"></param>
    ''' <param name="relative"></param>
    Public Sub MoveCameraXYZ(position As Vector3, relative As Boolean)
        If relative Then
            cameraPosition += position  '現在位置に加算
        Else
            cameraPosition = position   '現在位置を置換
        End If

        '行列を再計算
        cameraMatrix = Matrix4.CreateTranslation(cameraPosition)
        cameraMatrix = cameraMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitY, cameraHorizAngle)
        cameraMatrix = cameraMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitX, cameraVertAngle)
    End Sub

    ''' <summary>
    ''' カメラの向き（視線）を変更する
    ''' horizAngle: 水平方向の角度
    ''' vertAngle: 上下方向の角度
    ''' relative: Trueで現在角度に加算、Falseで現在角度に置き換える
    ''' </summary>
    ''' <param name="horizAngle"></param>
    ''' <param name="vertAngle"></param>
    ''' <param name="relative"></param>
    Public Sub TurnCamera(horizAngle As Single, vertAngle As Single, relative As Boolean)
        If relative Then
            '角度を追加
            cameraHorizAngle += horizAngle
            cameraVertAngle += vertAngle
        Else
            '角度を追加
            cameraHorizAngle = horizAngle
            cameraVertAngle = vertAngle
        End If

        '行列を再計算
        cameraMatrix = Matrix4.CreateTranslation(cameraPosition)
        cameraMatrix = cameraMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitY, cameraHorizAngle)
        cameraMatrix = cameraMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitX, cameraVertAngle)
    End Sub

    ''' <summary>
    ''' カメラの位置（視点）を変更する
    ''' distance: 水平方向、上下方向、前後方向の距離
    ''' </summary>
    ''' <param name="distance"></param>
    Public Sub MoveCamera(distance As Vector3)
        Dim m As Matrix4 = Matrix4.CreateFromAxisAngle(Vector3.UnitX, -cameraVertAngle)
        m = m * Matrix4.CreateFromAxisAngle(Vector3.UnitY, -cameraHorizAngle) 'm=カメラの向きを表す行列
        m = Matrix4.CreateTranslation(distance) * m 'パラメータの向きをカメラの向きに変換する
        cameraPosition += m.ExtractTranslation() '現在位置に加算

        '行列を再計算
        cameraMatrix = Matrix4.CreateTranslation(cameraPosition)
        cameraMatrix = cameraMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitY, cameraHorizAngle)
        cameraMatrix = cameraMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitX, cameraVertAngle)
    End Sub

    Private Sub glControl_MouseDown(sender As Object, e As MouseEventArgs)
        mousePos = New Point(e.X, e.Y)
    End Sub

    Private Sub glControl_MouseMove(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            '視線移動
            TurnCamera(CSng(e.X - mousePos.X) / CSng(Me.Width) * 2 * Math.PI,
                CSng(e.Y - mousePos.Y) / CSng(Me.Height) * 2 * Math.PI, True)
            mousePos = New Point(e.X, e.Y)
        ElseIf e.Button = MouseButtons.Left Then
            '視点移動
            MoveCamera(New Vector3(CSng(e.X - mousePos.X) * 0.005F * cameraDistance, -CSng(e.Y - mousePos.Y) * 0.005F * cameraDistance, 0.0F))
            mousePos = New Point(e.X, e.Y)
        End If
    End Sub

    Private Sub glControl_MouseWheel(sender As Object, e As MouseEventArgs)

        Dim yang As Single = (e.Y / Me.Height - 0.5F) * viewingAngleV
        Dim xang As Single = (e.X / Me.Width - 0.5F) * viewingAngleV * Me.Width / Me.Height

        '視点移動
        Dim m As Matrix4 = Matrix4.CreateFromAxisAngle(Vector3.UnitX, -cameraVertAngle - yang)
        m = m * Matrix4.CreateFromAxisAngle(Vector3.UnitY, -cameraHorizAngle - xang)
        m = Matrix4.CreateTranslation(0.0F, 0.0F, e.Delta / 10.0F) * m

        MoveCameraXYZ(m.ExtractTranslation(), relative:=True)

        'eyePosition += m.ExtractTranslation()


        'eyeMatrix = Matrix4.CreateTranslation(eyePosition)
        'eyeMatrix = eyeMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitY, modelHorizAngle)
        'eyeMatrix = eyeMatrix * Matrix4.CreateFromAxisAngle(Vector3.UnitX, modelVertAngle)
    End Sub

    ''' <summary>
    ''' 登録されているBufferObjectの一覧を取得する
    ''' </summary>
    ''' <returns></returns>
    Public Function GetBufferObjects() As BufferObject()
        Return arBO.ToArray()
    End Function

    ''' <summary>
    ''' 名前からBufferObjectを取得する
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    Public Function GetBufferObject(name As String) As BufferObject
        For Each bo As BufferObject In arBO
            If bo.GetName = name Then Return bo
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' 登録されているテクスチャの一覧を取得する
    ''' </summary>
    ''' <returns></returns>
    Public Function GetTextures() As TextureInfo()
        Return arTxt.ToArray()
    End Function

    ''' <summary>
    ''' 名前からテクスチャビットマップを取得する
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    Public Function GetTexture(name As String) As Bitmap
        Dim ti As TextureInfo
        For i As Integer = 0 To arTxt.Count - 1
            ti = arTxt(i)
            If ti.name = name Then Return ti.bmp
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' バッファオブジェクトによって割り当てられた頂点バッファなどを解放し、
    ''' 管理用配列からバッファオブジェクトを除外します。
    ''' テクスチャは別途破棄する必要があります。
    ''' アプリケーションの終了時にわざわざ呼び出す必要はありません。
    ''' </summary>
    ''' <param name="bo"></param>
    ''' <returns></returns>
    Public Function DeleteBufferObject(ByRef bo As BufferObject) As Boolean
        If bo Is Nothing Then Return False

        '頂点バッファが登録済みであれば削除する
        If bo._vao <> 0 Then
            GL.DeleteVertexArray(bo._vao)
            bo._vao = 0
        End If
        If bo._vbo <> 0 Then
            GL.DeleteBuffers(1, bo._vbo)
            bo._vbo = 0
        End If
        If bo._ibo <> 0 Then
            GL.DeleteBuffers(1, bo._ibo)
            bo._ibo = 0
            bo._numIndex = 0
        End If

        '管理配列から削除する
        arBO.Remove(bo)
        bo._parent = Nothing
        Return True
    End Function

    ''' <summary>
    ''' バッファオブジェクトによって割り当てられた頂点バッファなどを解放し、
    ''' 管理用配列からバッファオブジェクトを除外します。
    ''' テクスチャは別途破棄する必要があります。
    ''' アプリケーションの終了時にわざわざ呼び出す必要はありません。
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    Public Function DeleteBufferObject(name As String) As Boolean
        Dim bo As BufferObject = GetBufferObject(name)
        If bo Is Nothing Then Return False
        Return DeleteBufferObject(bo)
    End Function

    ''' <summary>
    ''' SetTexture()によって割り当てられたテクスチャバッファを削除し、
    ''' 管理用配列からも除外します。
    ''' bmpにはAddTexture()呼び出し時に使用したビットマップを渡します。
    ''' アプリケーションの終了時にわざわざ呼び出す必要はありません。
    ''' </summary>
    ''' <param name="bmp"></param>
    ''' <returns></returns>
    Public Function DeleteTexture(bmp As Bitmap) As Boolean
        For i As Integer = 0 To arTxt.Count - 1
            Dim ti As TextureInfo = arTxt(i)
            If ti.bmp Is bmp Then
                If ti.txt > 0 Then
                    GL.DeleteTexture(ti.txt)
                    ti.txt = 0
                End If
                arTxt.RemoveAt(i)
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' SetTexture()によって割り当てられたテクスチャバッファを削除し、管理用配列からも除外します。
    ''' bmpにはAddTexture()呼び出し時に設定した名前を渡します。
    ''' アプリケーションの終了時にわざわざ呼び出す必要はありません。
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    Public Function DeleteTexture(name As String) As Boolean
        For i As Integer = 0 To arTxt.Count - 1
            Dim ti As TextureInfo = arTxt(i)
            If ti.name <> "" AndAlso ti.name = name Then
                GL.DeleteTexture(ti.txt)
                arTxt.RemoveAt(i)
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' ビットマップをテクスチャバッファに追加します。
    ''' テクスチャに名前をつけることができます。名前に使える文字はファイル名に使える文字のみです。
    ''' テクスチャへの参照方法はビットマップまたは名前を使うことができます。
    ''' ビットマップや、設定した名前が重複していた場合、もしくは名前に不適正な文字があった場合はFalseを返して追加されません。
    ''' </summary>
    ''' <param name="bmp"></param>
    ''' <param name="name"></param>
    ''' <return></return>
    Public Function AddTexture(bmp As Bitmap, Optional name As String = "") As Boolean
        If name.IndexOfAny(IO.Path.GetInvalidFileNameChars()) >= 0 Then Return False

        Dim newTexture As Boolean = True

        For i As Integer = 0 To arTxt.Count - 1
            With arTxt(i)
                If .bmp Is bmp Then Return False 'ビットマップが重複
                If .name <> "" AndAlso name <> "" AndAlso .name = name Then Return False '名前が重複
            End With
        Next

        Dim ti As New TextureInfo
        ti.bmp = bmp
        ti.name = name

        'Textureの許可
        GL.Enable(EnableCap.Texture2D)

        'テクスチャ用バッファの生成
        ti.txt = GL.GenTexture()

        'テクスチャ用バッファのひもづけ
        GL.BindTexture(TextureTarget.Texture2D, ti.txt)

        'テクスチャの設定
        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMinFilter, TextureMinFilter.Nearest)
        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMagFilter, TextureMagFilter.Nearest)

        'データ読み込み
        Dim Data As Imaging.BitmapData = bmp.LockBits(New Rectangle(0, 0, bmp.Width, bmp.Height), Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)

        'テクスチャ用バッファに色情報を流し込む
        GL.TexImage2D(TextureTarget.Texture2D, 0, PixelInternalFormat.Rgba, Data.Width, Data.Height, 0, OpenTK.Graphics.OpenGL.PixelFormat.Bgra, PixelType.UnsignedByte, Data.Scan0)

        bmp.UnlockBits(Data)
        GL.BindTexture(TextureTarget.Texture2D, 0)

        'テクスチャ管理配列に追加
        If newTexture Then arTxt.Add(ti)
        Return True
    End Function

    Friend Function FindTexture(bmp As Bitmap) As TextureInfo
        Dim ti As TextureInfo
        For i As Integer = 0 To arTxt.Count - 1
            ti = arTxt(i)
            If ti.bmp Is bmp Then Return ti
        Next
        Return Nothing
    End Function

    Friend Function FindTexture(name As String) As TextureInfo
        Dim ti As TextureInfo
        For i As Integer = 0 To arTxt.Count - 1
            ti = arTxt(i)
            If ti.name = name Then Return ti
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' 登録済みのテクスチャを更新します
    ''' 名前で指定する場合はnameパラメータを設定してください
    ''' テクスチャが見つからない場合はFalseを返し更新されません
    ''' </summary>
    ''' <param name="bmp"></param>
    ''' <param name="name"></param>
    ''' <returns></returns>
    Public Function UpdateTexture(bmp As Bitmap, Optional name As String = "") As Boolean
        Dim found As Boolean = False
        Dim ti As TextureInfo = Nothing

        For i As Integer = 0 To arTxt.Count - 1
            ti = arTxt(i)
            If (name = "" AndAlso ti.bmp Is bmp) OrElse (name <> "" AndAlso ti.name = name) Then
                If ti.txt > 0 Then
                    GL.DeleteTexture(ti.txt)
                    ti.txt = 0
                End If
                found = True
                Exit For
            End If
        Next

        If Not found Then Return False '見つからない

        'Textureの許可
        GL.Enable(EnableCap.Texture2D)

        'テクスチャ用バッファの生成
        ti.txt = GL.GenTexture()

        'テクスチャ用バッファのひもづけ
        GL.BindTexture(TextureTarget.Texture2D, ti.txt)

        'テクスチャの設定
        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMinFilter, TextureMinFilter.Nearest)
        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMagFilter, TextureMagFilter.Nearest)

        'データ読み込み
        Dim Data As Imaging.BitmapData = bmp.LockBits(New Rectangle(0, 0, bmp.Width, bmp.Height), Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)

        'テクスチャ用バッファに色情報を流し込む
        GL.TexImage2D(TextureTarget.Texture2D, 0, PixelInternalFormat.Rgba, Data.Width, Data.Height, 0, OpenTK.Graphics.OpenGL.PixelFormat.Bgra, PixelType.UnsignedByte, Data.Scan0)

        bmp.UnlockBits(Data)
        GL.BindTexture(TextureTarget.Texture2D, 0)

        Return True
    End Function

    ''' <summary>
    ''' Z座標における、XY平面の表示エリアの境界線をRectangleFで返す
    ''' 2Dモデルを画面の端に表示したい場合、この関数の返す境界を使って座標を決定することができる
    ''' この関数はビューポートが変更されると返す値が変化するため、通常は ViewportResized ハンドラを使って再計算、再配置する
    ''' </summary>
    ''' <param name="zValue"></param>
    ''' <returns></returns>
    Public Function GetVisibleEdgeFor2D(Optional zValue As Single = 0.0F) As RectangleF
        Dim a As Single = cameraDistance - zValue
        Dim c As Single = a / Math.Cos(viewingAngleV / 2.0F)
        Dim b As Single = Math.Sqrt(c * c - a * a)
        Dim d As Single = CSng(Me.Width) / CSng(Me.Height) * b
        Return New RectangleF(-d, b, d * 2.0F, -b * 2.0F)
    End Function

    ''' <summary>
    ''' 3Dオブジェクトを構築するためのBufferObjectを返す
    ''' この関数が実行された時点で、返されたBufferObjectは描画対象としてView3D内の管理配列に保存されます
    ''' View3Dの管理対象から削除するにはDeleteBufferObject()を呼び出します
    ''' パラメータは頂点バッファのデータ構成を指定します
    ''' bNormal: 法線ベクトルを使う（照明による陰影処理に必要）
    ''' bTexture: テクスチャ座標を持つ（テクスチャを貼り付ける場合に必要）
    ''' bColor: 頂点カラーを持つ（テクスチャやマテリアルを使わずに色をつける場合）
    ''' </summary>
    ''' <param name="bNormal"></param>
    ''' <param name="bTexture"></param>
    ''' <param name="bColor"></param>
    ''' <returns></returns>
    Public Function Create3DObject(bNormal As Boolean, bTexture As Boolean, bColor As Boolean, Optional name As String = "") As BufferObject
        If name.IndexOfAny(IO.Path.GetInvalidFileNameChars()) >= 0 Then Return Nothing
        arBO.Add(New BufferObject(Me, True, bNormal, bTexture, bColor, False, name))
        Return arBO.Last
    End Function

    ''' <summary>
    ''' カメラの位置や向きに依存しない、2Dオブジェクトを構築するためのBufferObjectを返す
    ''' この関数が実行された時点で、返されたBufferObjectは描画対象としてView3D内の管理配列に保存されます
    ''' View3Dの管理対象から削除するにはDeleteBufferObject()を呼び出します
    ''' パラメータは頂点バッファのデータ構成を指定します
    ''' bNormal: 法線ベクトルを使う（照明による陰影処理に必要）
    ''' bTexture: テクスチャ座標を持つ（テクスチャを貼り付ける場合に必要）
    ''' bColor: 頂点カラーを持つ（テクスチャやマテリアルを使わずに色をつける場合）
    ''' </summary>
    ''' <param name="bNormal"></param>
    ''' <param name="bTexture"></param>
    ''' <param name="bColor"></param>
    ''' <returns></returns>
    Public Function Create2DObject(bNormal As Boolean, bTexture As Boolean, bColor As Boolean, Optional name As String = "") As BufferObject
        If name.IndexOfAny(IO.Path.GetInvalidFileNameChars()) >= 0 Then Return Nothing
        arBO.Add(New BufferObject(Me, False, bNormal, bTexture, bColor, False, name))
        Return arBO.Last
    End Function

    ''' <summary>
    ''' ポイントスプライトに対応したオブジェクトを構築するためのBufferObjectを返す
    ''' ポイントスプライトとは、3Dオブジェクトながら常にこちらを向いたテクスチャを表示するようなものです
    ''' この関数が実行された時点で、返されたBufferObjectは描画対象としてView3D内の管理配列に保存されます
    ''' View3Dの管理対象から削除するにはDeleteBufferObject()を呼び出します
    ''' パラメータは頂点バッファのデータ構成を指定します
    ''' bTexture: テクスチャ座標を持つ（テクスチャを貼り付ける場合に必要）
    ''' bColor: 頂点カラーを持つ（テクスチャやマテリアルを使わずに色をつける場合）
    ''' </summary>
    ''' <param name="bTexture"></param>
    ''' <param name="bColor"></param>
    ''' <returns></returns>
    Public Function CreateProintSprite(bTexture As Boolean, bColor As Boolean, Optional name As String = "") As BufferObject
        If name.IndexOfAny(IO.Path.GetInvalidFileNameChars()) >= 0 Then Return Nothing
        arBO.Add(New BufferObject(Me, True, False, bTexture, bColor, True, name))
        Return arBO.Last
    End Function

    ''' <summary>
    ''' DeleteBufferObjectで削除されたBufferObjectを再登録する場合に使います
    ''' 既に登録されている場合はFalseを返します
    ''' </summary>
    ''' <param name="bo"></param>
    ''' <returns></returns>
    Public Function AddBufferObject(bo As BufferObject) As Boolean
        If arBO.IndexOf(bo) >= 0 Then Return False '登録済み
        arBO.Add(bo)
        bo._parent = Me
        Return True
    End Function


    'XML出力用の補助クラス
    Friend Class XmlWriteNode
        Public Name As String
        Public Value As Object  'String または List(Of XmlNode)
        Public Sub New(name As String, value As String)
            Me.Name = name
            Me.Value = value
        End Sub
        Public Sub New(name As String)
            Me.Name = name
            Me.Value = New List(Of XmlWriteNode)
        End Sub
        Public Sub Add(node As XmlWriteNode)
            Me.Value.add(node)
        End Sub
        Public Sub Add(name As String, value As String)
            Me.Value.add(New XmlWriteNode(name, value))
        End Sub
        Public Sub Write(xmlWtr As XmlTextWriter, Optional indentCount As Integer = 0)
            If TypeOf Value Is String Then
                If indentCount > 0 Then xmlWtr.WriteString(New String(" "c, indentCount * 4))
                xmlWtr.WriteStartElement(Me.Name)
                xmlWtr.WriteString(Value)
                xmlWtr.WriteEndElement()
                xmlWtr.WriteString(Chr(13) & Chr(10))
            Else
                If indentCount > 0 Then xmlWtr.WriteString(New String(" "c, indentCount * 4))
                xmlWtr.WriteStartElement(Me.Name)
                xmlWtr.WriteString(Chr(13) & Chr(10))
                For Each node In Value
                    node.Write(xmlWtr, indentCount + 1)
                Next
                If indentCount > 0 Then xmlWtr.WriteString(New String(" "c, indentCount * 4))
                xmlWtr.WriteEndElement()
                xmlWtr.WriteString(Chr(13) & Chr(10))
            End If
        End Sub
    End Class


    ''' <summary>
    ''' ３D環境をXMLファイルに保存します。
    ''' 保存する項目はパラメータで指定します
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <param name="bView"></param>
    ''' <param name="bCamera"></param>
    ''' <param name="bLight"></param>
    ''' <param name="bOthers"></param>
    ''' <returns></returns>
    Public Function SaveToXml(filePath As String, bView As Boolean, bCamera As Boolean, bLight As Boolean, bOthers As Boolean) As Boolean

        Try
            Dim xmlWtr As New XmlTextWriter(filePath, System.Text.Encoding.UTF8)

            'XML宣言を出力
            xmlWtr.WriteStartDocument()
            xmlWtr.WriteStartElement("Environments")
            xmlWtr.WriteString(Chr(13) & Chr(10))

            Dim node As XmlWriteNode
            If bView Then
                node = New XmlWriteNode("View")
                node.Add("viewingAngleV", viewingAngleV.ToString)
                node.Add("zNear", zNear.ToString)
                node.Add("zFar", zFar.ToString)
                node.Write(xmlWtr, 1)
            End If

            If bCamera Then
                node = New XmlWriteNode("Camera")
                node.Add("cameraHorizAngle", cameraHorizAngle.ToString)
                node.Add("cameraVertAngle", cameraVertAngle.ToString)
                node.Add("cameraPosition", cameraPosition.ToString)
                node.Add("cameraDistance", cameraDistance.ToString)
                node.Write(xmlWtr, 1)
            End If

            If bLight Then
                node = New XmlWriteNode("Lights")
                For i As Integer = 0 To 3
                    If _light(i) IsNot Nothing Then
                        With _light(i)
                            Dim nodeLight As New XmlWriteNode("Light")
                            nodeLight.Add("Index", i.ToString)
                            nodeLight.Add("Active", .Active.ToString)
                            nodeLight.Add("Position", .Position.ToString)
                            nodeLight.Add("Diffuse", .Diffuse.ToString)
                            nodeLight.Add("Specular", .Specular.ToString)
                            nodeLight.Add("Ambient", .Ambient.ToString)
                            nodeLight.Add("FollowCamera", .FollowCamera.ToString)
                            node.Add(nodeLight)
                        End With
                    End If
                Next
                node.Write(xmlWtr, 1)
            End If

            If bOthers Then
                node = New XmlWriteNode("Others")
                node.Add("_clearColor", _clearColor.ToString)
                node.Add("_limitFrames", _limitFrames.ToString)
                node.Write(xmlWtr, 1)
            End If

            xmlWtr.WriteEndElement()
            xmlWtr.WriteEndDocument()
            xmlWtr.Flush()
            xmlWtr.Close()
        Catch ex As Exception
            Debug.Print("View3Dの環境保存に失敗しました")
            Debug.Print(ex.Message)
            Return False
        End Try

        Return True
    End Function

    '文字列"(a,b,c,d)"から配列{"a","b","c","d"}を得る
    Friend Shared Function GetValuesInBrackets(s As String) As String()
        Dim i As Integer = s.IndexOf("("c)
        Dim j As Integer = s.IndexOf(")"c, i)
        Return s.Substring(i + 1, j - i - 1).Split({","c})
    End Function

    '文字列からColor4に変換
    '文字列の例 {(R, G, B, A) = (0.7, 0.7, 0.7, 1)}
    Friend Shared Function ToColor4(s As String) As Color4
        Dim ar = s.Split({"="c})
        Dim arTag = GetValuesInBrackets(ar(0))
        Dim arValue = GetValuesInBrackets(ar(1))
        Dim c As New Color4
        For i As Integer = 0 To UBound(arTag)
            Select Case arTag(i).Trim
                Case "R"c : c.R = Single.Parse(arValue(i))
                Case "G"c : c.G = Single.Parse(arValue(i))
                Case "B"c : c.B = Single.Parse(arValue(i))
                Case "A"c : c.A = Single.Parse(arValue(i))
            End Select
        Next
        Return c
    End Function

    '文字列からColor4に変換
    '文字列の例 {(10, 0, 0, 0),(0, 10, 0, 0),(0, 0, 10, 0),(0, 0, 0, 1)}
    Friend Shared Function ToMatrix4(s As String) As Matrix4
        Dim ar = s.Split({",("}, StringSplitOptions.None)
        Dim m As New Matrix4
        For i As Integer = 0 To UBound(ar)
            Dim ar2 = GetValuesInBrackets(IIf(i > 0, "(", "") & ar(i))
            Select Case i
                Case 0
                    m.M11 = Single.Parse(ar2(0))
                    m.M12 = Single.Parse(ar2(1))
                    m.M13 = Single.Parse(ar2(2))
                    m.M14 = Single.Parse(ar2(3))
                Case 1
                    m.M21 = Single.Parse(ar2(0))
                    m.M22 = Single.Parse(ar2(1))
                    m.M23 = Single.Parse(ar2(2))
                    m.M24 = Single.Parse(ar2(3))
                Case 2
                    m.M31 = Single.Parse(ar2(0))
                    m.M32 = Single.Parse(ar2(1))
                    m.M33 = Single.Parse(ar2(2))
                    m.M34 = Single.Parse(ar2(3))
                Case 3
                    m.M41 = Single.Parse(ar2(0))
                    m.M42 = Single.Parse(ar2(1))
                    m.M43 = Single.Parse(ar2(2))
                    m.M44 = Single.Parse(ar2(3))
            End Select
        Next
        Return m
    End Function

    '文字列からVector3に変換
    '文字列の例 (1, 2, 3)
    Friend Shared Function ToVector3(s As String) As Vector3
        Dim ar = GetValuesInBrackets(s)
        Dim v As New Vector3
        v.x = Single.Parse(ar(0))
        v.Y = Single.Parse(ar(1))
        v.Z = Single.Parse(ar(2))
        Return v
    End Function

    '文字列からVector4に変換
    '文字列の例 (1, 2, 3, 4)
    Friend Shared Function ToVector4(s As String) As Vector4
        Dim ar = GetValuesInBrackets(s)
        Dim v As New Vector4
        v.X = Single.Parse(ar(0))
        v.Y = Single.Parse(ar(1))
        v.Z = Single.Parse(ar(2))
        v.W = Single.Parse(ar(3))
        Return v
    End Function

    ''' <summary>
    ''' SaveToXml()で保存した３D環境を読み込みます。
    ''' 読み込む項目はパラメータで指定します
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <param name="bView"></param>
    ''' <param name="bCamera"></param>
    ''' <param name="bLight"></param>
    ''' <param name="bOthers"></param>
    ''' <returns></returns>
    Public Function LoadFromXml(filePath As String, bView As Boolean, bCamera As Boolean, bLight As Boolean, bOthers As Boolean) As Boolean

        Try
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(filePath)

            Dim Environments As XmlNode = xmlDoc.GetElementsByTagName("Environments").Item(0)
            Dim n As XmlNode

            If bView Then
                n = Environments.SelectSingleNode("View")
                If n IsNot Nothing Then
                    viewingAngleV = Single.Parse(n.SelectSingleNode("viewingAngleV").InnerText)
                    zNear = Single.Parse(n.SelectSingleNode("zNear").InnerText)
                    zFar = Single.Parse(n.SelectSingleNode("zFar").InnerText)
                    PrepareViewport()
                End If
            End If

            If bCamera Then
                n = Environments.SelectSingleNode("Camera")
                If n IsNot Nothing Then
                    Dim horizAngle As Single = Single.Parse(n.SelectSingleNode("cameraHorizAngle").InnerText)
                    Dim vertAngle As Single = Single.Parse(n.SelectSingleNode("cameraVertAngle").InnerText)
                    TurnCamera(horizAngle, vertAngle, False)

                    Dim position As Vector3 = ToVector3(n.SelectSingleNode("cameraPosition").InnerText)
                    MoveCameraXYZ(position, relative:=False)

                    cameraDistance = Single.Parse(n.SelectSingleNode("cameraDistance").InnerText)
                End If
            End If

            If bLight Then
                n = Environments.SelectSingleNode("Lights")
                If n IsNot Nothing Then
                    For Each ln As XmlNode In n.SelectNodes("Light")
                        Dim light As New Light
                        Dim index As Integer = Integer.Parse(ln.SelectSingleNode("Index").InnerText)
                        _light(index) = light
                        light.Active = Boolean.Parse(ln.SelectSingleNode("Active").InnerText)
                        light.Position = ToVector4(ln.SelectSingleNode("Position").InnerText)
                        light.Diffuse = ToColor4(ln.SelectSingleNode("Diffuse").InnerText)
                        light.Specular = ToColor4(ln.SelectSingleNode("Specular").InnerText)
                        light.Ambient = ToColor4(ln.SelectSingleNode("Ambient").InnerText)
                        light.FollowCamera = Boolean.Parse(ln.SelectSingleNode("FollowCamera").InnerText)
                    Next
                End If
            End If

            If bOthers Then
                n = Environments.SelectSingleNode("Others")
                If n IsNot Nothing Then
                    _clearColor = ToColor4(n.SelectSingleNode("_clearColor").InnerText)
                    _limitFrames = Single.Parse(n.SelectSingleNode("_limitFrames").InnerText)
                End If
            End If

        Catch ex As Exception
            Debug.Print("View3Dの環境読み込みに失敗しました")
            Debug.Print(ex.Message)
            Return False
        End Try

        Return True
    End Function



    '光源を管理
    Public Class Light

        '光源の基本パラメータ
        Public Position As Vector4
        Public Ambient As Color4
        Public Diffuse As Color4
        Public Specular As Color4

        'Trueにすると、光源をカメラの動きに追随させる（視野が動いても、視野に対して同じ位置に光源がある）
        'Falseは光源はグローバルな位置になる（デフォルト）
        Public FollowCamera As Boolean = False

        Public Active As Boolean = True

        ''' <summary>
        ''' 適当な平行光源が作成されます
        ''' </summary>
        Sub New()
            '適当
            Position = New Vector4(200.0F, 150.0F, 500.0F, 0.0F)
            Ambient = Color4.White
            Diffuse = New Color4(0.7F, 0.7F, 0.7F, 1.0F)
            Specular = New Color4(1.0F, 1.0F, 1.0F, 1.0F)
            Active = True
        End Sub

        ''' <summary>
        ''' 詳細を指定して光源を作成します
        ''' </summary>
        ''' <param name="position"></param>
        ''' <param name="ambient"></param>
        ''' <param name="diffuse"></param>
        ''' <param name="specular"></param>
        Public Sub New(position As Vector4, ambient As Color4, diffuse As Color4, specular As Color4)
            Me.Position = position
            Me.Ambient = ambient
            Me.Diffuse = diffuse
            Me.Specular = specular
            Active = True
        End Sub

        ''' <summary>
        ''' 平行光源を作成します。平行光源とは太陽のような位置に依らず、平行に光を発する光源です
        ''' direction: 光の向き
        ''' color: 色
        ''' ambient: 環境光の強さ 0.0F～1.0F
        ''' </summary>
        ''' <param name="direction"></param>
        ''' <param name="color"></param>
        ''' <param name="ambient"></param>
        ''' <returns></returns>
        Public Shared Function CreateParallelLight(direction As Vector3, color As Color4, ambient As Single) As Light
            Return New Light(
                New Vector4(direction, 0.0F),
                New Color4(ambient * color.R, ambient * color.G, ambient * color.B, 1.0F),
                color,
                color)
        End Function

        ''' <summary>
        ''' 点光源を作成します。点光源とは電球のようにある位置から全方向を照らす光源です
        ''' position: 電球の位置
        ''' color: 色
        ''' ambient: 環境光の強さ 0.0F～1.0F
        ''' </summary>
        ''' <param name="position"></param>
        ''' <param name="color"></param>
        ''' <param name="ambient"></param>
        ''' <returns></returns>
        Public Shared Function CreatePointLight(position As Vector3, color As Color4, ambient As Single) As Light
            Return New Light(
                New Vector4(position, 1.0F),
                New Color4(ambient * color.R, ambient * color.G, ambient * color.B, ambient * color.A),
                color,
                color)
        End Function
    End Class

    '材質を管理
    Public Class Material
        'http://www.natural-science.or.jp/article/20110313211402.php
        Public Enum Preset
            Ruby            'ruby(ルビー)
            Emerald         'emerald(エメラルド)
            Jade            'jade(翡翠)
            Obsidian        'obsidian(黒曜石)
            Pearl           'pearl(真珠)
            Turquoise       'turquoise(トルコ石)
            Brass           'brass(真鍮)
            Bronze          'bronze(青銅)
            Chrome          'chrome(クローム)
            Copper          'copper(銅)
            Gold            'gold(金)
            Silver          'silver(銀)
            PlasticBlack    'プラスチック(黒)
            PlasticCyan     'プラスチック(シアン)
            PlasticGreen    'プラスチック(緑)
            PlasticRed      'プラスチック(赤)
            PlasticWhite    'プラスチック(白)
            PlasticYellow   'プラスチック(黄)
            RubberBlack     'ゴム(黒)
            RubberCyan      'ゴム(シアン)
            RubberGreen     'ゴム(緑)
            RubberRed       'ゴム(赤)
            RubberWhite     'ゴム(白)
            RubberYellow    'ゴム(黄)
        End Enum

        Public Ambient As Color4
        Public Diffuse As Color4
        Public Specular As Color4
        Public Shininess As Single

        ''' <summary>
        ''' 適当なマテリアルを作成します
        ''' </summary>
        Public Sub New()
            '適当な材質
            Ambient = New Color4(0.2F, 0.2F, 0.2F, 1.0F)
            Diffuse = New Color4(0.7F, 0.7F, 0.7F, 1.0F)
            Specular = New Color4(0.6F, 0.6F, 0.6F, 1.0F)
            Shininess = 51.4F
        End Sub

        ''' <summary>
        ''' 詳細を指定してマテリアルを作成します
        ''' </summary>
        ''' <param name="ambient"></param>
        ''' <param name="diffuse"></param>
        ''' <param name="specular"></param>
        ''' <param name="shininess"></param>
        Public Sub New(ambient As Color4, diffuse As Color4, specular As Color4, shininess As Single)
            Me.Ambient = ambient
            Me.Diffuse = diffuse
            Me.Specular = specular
            Me.Shininess = shininess
        End Sub

        ''' <summary>
        ''' マテリアルをプリセットから作成します
        ''' </summary>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Public Shared Function FromPreset(type As Preset) As Material
            Dim m As New Material
            Select Case type
                Case Preset.Ruby 'ruby(ルビー)
                    m.Ambient = New Color4(0.1745F, 0.01175F, 0.01175F, 1.0F)
                    m.Diffuse = New Color4(0.61424F, 0.04136F, 0.04136F, 1.0F)
                    m.Specular = New Color4(0.727811F, 0.626959F, 0.626959F, 1.0F)
                    m.Shininess = 76.8F
                Case Preset.Emerald 'emerald(エメラルド)
                    m.Ambient = New Color4(0.0215F, 0.1745F, 0.0215F, 1.0F)
                    m.Diffuse = New Color4(0.07568F, 0.61424F, 0.07568F, 1.0F)
                    m.Specular = New Color4(0.633F, 0.727811F, 0.633F, 1.0F)
                    m.Shininess = 76.8F
                Case Preset.Jade 'jade(翡翠)
                    m.Ambient = New Color4(0.135F, 0.2225F, 0.1575F, 1.0F)
                    m.Diffuse = New Color4(0.54F, 0.89F, 0.63F, 1.0F)
                    m.Specular = New Color4(0.316228F, 0.316228F, 0.316228F, 1.0F)
                    m.Shininess = 12.8F
                Case Preset.Obsidian 'obsidian(黒曜石)
                    m.Ambient = New Color4(0.05375F, 0.05F, 0.06625F, 1.0F)
                    m.Diffuse = New Color4(0.18275F, 0.17F, 0.22525F, 1.0F)
                    m.Specular = New Color4(0.332741F, 0.328634F, 0.346435F, 1.0F)
                    m.Shininess = 38.4F
                Case Preset.Pearl ' pearl(真珠)
                    m.Ambient = New Color4(0.25F, 0.20725F, 0.20725F, 1.0F)
                    m.Diffuse = New Color4(1.0F, 0.829F, 0.829F, 1.0F)
                    m.Specular = New Color4(0.296648F, 0.296648F, 0.296648F, 1.0F)
                    m.Shininess = 10.24F
                Case Preset.Turquoise 'turquoise(トルコ石)
                    m.Ambient = New Color4(0.1F, 0.18725F, 0.1745F, 1.0F)
                    m.Diffuse = New Color4(0.396F, 0.74151F, 0.69102F, 1.0F)
                    m.Specular = New Color4(0.297254F, 0.30829F, 0.306678F, 1.0F)
                    m.Shininess = 12.8F
                Case Preset.Brass 'brass(真鍮)
                    m.Ambient = New Color4(0.329412F, 0.223529F, 0.027451F, 1.0F)
                    m.Diffuse = New Color4(0.780392F, 0.568627F, 0.113725F, 1.0F)
                    m.Specular = New Color4(0.992157F, 0.941176F, 0.807843F, 1.0F)
                    m.Shininess = 27.8974361F
                Case Preset.Bronze 'bronze(青銅)
                    m.Ambient = New Color4(0.2125F, 0.1275F, 0.054F, 1.0F)
                    m.Diffuse = New Color4(0.714F, 0.4284F, 0.18144F, 1.0F)
                    m.Specular = New Color4(0.393548F, 0.271906F, 0.166721F, 1.0F)
                    m.Shininess = 25.6F
                Case Preset.Chrome 'chrome(クローム)
                    m.Ambient = New Color4(0.25F, 0.25F, 0.25F, 1.0F)
                    m.Diffuse = New Color4(0.4F, 0.4F, 0.4F, 1.0F)
                    m.Specular = New Color4(0.774597F, 0.774597F, 0.774597F, 1.0F)
                    m.Shininess = 76.8F
                Case Preset.Copper 'copper(銅)
                    m.Ambient = New Color4(0.19125F, 0.0735F, 0.0225F, 1.0F)
                    m.Diffuse = New Color4(0.7038F, 0.27048F, 0.0828F, 1.0F)
                    m.Specular = New Color4(0.256777F, 0.137622F, 0.086014F, 1.0F)
                    m.Shininess = 12.8F
                Case Preset.Gold 'gold(金)
                    m.Ambient = New Color4(0.24725F, 0.1995F, 0.0745F, 1.0F)
                    m.Diffuse = New Color4(0.75164F, 0.60648F, 0.22648F, 1.0F)
                    m.Specular = New Color4(0.628281F, 0.555802F, 0.366065F, 1.0F)
                    m.Shininess = 51.2F
                Case Preset.Silver 'silver(銀)
                    m.Ambient = New Color4(0.19225F, 0.19225F, 0.19225F, 1.0F)
                    m.Diffuse = New Color4(0.50754F, 0.50754F, 0.50754F, 1.0F)
                    m.Specular = New Color4(0.508273F, 0.508273F, 0.508273F, 1.0F)
                    m.Shininess = 51.2F
                Case Preset.PlasticBlack 'プラスチック(黒)
                    m.Ambient = New Color4(0.0F, 0.0F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.01F, 0.01F, 0.01F, 1.0F)
                    m.Specular = New Color4(0.5F, 0.5F, 0.5F, 1.0F)
                    m.Shininess = 32.0F
                Case Preset.PlasticCyan 'プラスチック(シアン)
                    m.Ambient = New Color4(0.0F, 0.1F, 0.06F, 1.0F)
                    m.Diffuse = New Color4(0.0F, 0.5098039F, 0.5098039F, 1.0F)
                    m.Specular = New Color4(0.501960754F, 0.501960754F, 0.501960754F, 1.0F)
                    m.Shininess = 32.0F
                Case Preset.PlasticGreen 'プラスチック(緑)
                    m.Ambient = New Color4(0.0F, 0.0F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.1F, 0.35F, 0.1F, 1.0F)
                    m.Specular = New Color4(0.45F, 0.55F, 0.45F, 1.0F)
                    m.Shininess = 32.0F
                Case Preset.PlasticRed 'プラスチック(赤)
                    m.Ambient = New Color4(0.0F, 0.0F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.5F, 0.0F, 0.0F, 1.0F)
                    m.Specular = New Color4(0.7F, 0.6F, 0.6F, 1.0F)
                    m.Shininess = 32.0F
                Case Preset.PlasticWhite 'プラスチック(白)
                    m.Ambient = New Color4(0.0F, 0.0F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.55F, 0.55F, 0.55F, 1.0F)
                    m.Specular = New Color4(0.7F, 0.7F, 0.7F, 1.0F)
                    m.Shininess = 32.0F
                Case Preset.PlasticYellow 'プラスチック(黄)
                    m.Ambient = New Color4(0.0F, 0.0F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.5F, 0.5F, 0.0F, 1.0F)
                    m.Specular = New Color4(0.6F, 0.6F, 0.5F, 1.0F)
                    m.Shininess = 32.0F
                Case Preset.RubberBlack 'ゴム(黒)
                    m.Ambient = New Color4(0.02F, 0.02F, 0.02F, 1.0F)
                    m.Diffuse = New Color4(0.01F, 0.01F, 0.01F, 1.0F)
                    m.Specular = New Color4(0.4F, 0.4F, 0.4F, 1.0F)
                    m.Shininess = 10.0F
                Case Preset.RubberCyan 'ゴム(シアン)
                    m.Ambient = New Color4(0.0F, 0.05F, 0.05F, 1.0F)
                    m.Diffuse = New Color4(0.4F, 0.5F, 0.5F, 1.0F)
                    m.Specular = New Color4(0.04F, 0.7F, 0.7F, 1.0F)
                    m.Shininess = 10.0F
                Case Preset.RubberGreen 'ゴム(緑)
                    m.Ambient = New Color4(0.0F, 0.05F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.4F, 0.5F, 0.4F, 1.0F)
                    m.Specular = New Color4(0.04F, 0.7F, 0.04F, 1.0F)
                    m.Shininess = 10.0F
                Case Preset.RubberRed 'ゴム(赤)
                    m.Ambient = New Color4(0.05F, 0.0F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.5F, 0.4F, 0.4F, 1.0F)
                    m.Specular = New Color4(0.7F, 0.04F, 0.04F, 1.0F)
                    m.Shininess = 10.0F
                Case Preset.RubberWhite 'ゴム(白)
                    m.Ambient = New Color4(0.05F, 0.05F, 0.05F, 1.0F)
                    m.Diffuse = New Color4(0.5F, 0.5F, 0.5F, 1.0F)
                    m.Specular = New Color4(0.7F, 0.7F, 0.7F, 1.0F)
                    m.Shininess = 10.0F
                Case Preset.RubberYellow 'ゴム(黄)
                    m.Ambient = New Color4(0.05F, 0.05F, 0.0F, 1.0F)
                    m.Diffuse = New Color4(0.5F, 0.5F, 0.4F, 1.0F)
                    m.Specular = New Color4(0.7F, 0.7F, 0.04F, 1.0F)
                    m.Shininess = 10.0F
            End Select
            Return m
        End Function
    End Class


    'View3Dで表示するオブジェクトを管理
    Public Class BufferObject
        Private name As String
        Public Material As View3D.Material = Nothing 'マテリアル(質感)を設定できる
        Public Lighting As Boolean = True       'ライティングを行うか
        Public Culling As Boolean = True        '裏面を消去するか
        Public Blend As Boolean = False         'アルファブレンディングを使用する
        Public PointSize As Single = 0.0F       'Point/Point sprite描画時のサイズ
        Public PointPerspective As Boolean = False 'Point sprite時に遠近感を持たせるか
        Public Active As Boolean = True         'オブジェクトを表示しない場合はFalseを設定する
        Public Tag As Object = Nothing          'ユーザーが使用できる汎用

        Public Martices As New List(Of Matrix4) '姿勢制御用配列。デフォルトで１つの配列を持っている。
        'ここに追加していくことで、同じオブジェクトを複数表示できる

        '_arb内の頂点データの構成(データの並び順は以下の変数順とする)
        Friend _bXYZ As Boolean                 'Vector3型の、 True:三次元座標(XYZ) False:二次元座標(XYW)
        Friend _bNormal As Boolean              'Vector3型の法線ベクトル
        Friend _bTexture As Boolean             'Vector2型のテクスチャ座標(0.0～1.0)
        Friend _bColor As Boolean               'Color4型の色

        Friend _bPointSprite As Boolean         'ポイントスプライト

        Private _primitive As BeginMode         'プリミティブを記憶
        Private _ari As New List(Of Int32)      '頂点インデックスを格納
        Private _arb As New List(Of Byte)       '頂点データを格納
        Private _numVertex As Integer = 0       '頂点の数  

        'テクスチャが設定された場合、下記のいずれかに値が入る（両方には入らない）
        Private _bmp As Bitmap = Nothing        'テクスチャ
        Private _textureName As String = ""     'テクスチャ名

        Friend _parent As View3D               'このクラスを生成した親View3Dクラスへの参照を保持

        '頂点バッファ関連
        Friend _vbo As Integer = 0
        Friend _ibo As Integer = 0
        Friend _vao As Integer = 0
        Friend _numIndex As Integer = 0
        Friend _bm As BeginMode

        Private Class Txt
            Public name As String = ""
            Public bmp As Bitmap = Nothing
            Public fileName As String = ""
            Public Sub New(name As String, bmp As Bitmap, fileName As String)
                Me.name = name
                Me.bmp = bmp
                Me.fileName = fileName
            End Sub
        End Class

        ''' <summary>
        ''' 複数のBufferObjectを１つの専用フォーマットファイル(XML形式)に保存します。
        ''' saveTexture=Trueでテクスチャファイルも保存します。BMP形式でtextureDirectoryディレクトリに保存されます。
        ''' テクスチャファイル名には"名前.bmp"で保存されますが、名前の無いテクスチャには"prefix001.bmp"形式で保存されます。
        ''' 各ファイルは常に上書き保存するので注意してください。
        ''' BufferObjectがView3Dクラスから削除されている場合は、テクスチャは保存されません。（検索できない）
        ''' savedFileListを渡すと、保存されたファイルのフルパスの一覧が返されます。
        ''' 成功したとき、関数はTrueを返します
        ''' </summary>
        ''' <param name="arBO"></param>
        ''' <param name="filePath"></param>
        ''' <param name="saveTexture"></param>
        ''' <param name="textureDirectory"></param>
        ''' <param name="prefix"></param>
        ''' <param name="savedFileList"></param>
        ''' <returns></returns>
        Public Shared Function SaveToXml(arBO As BufferObject(), filePath As String,
                                         saveTexture As Boolean, Optional textureDirectory As String = "",
                                         Optional prefix As String = "T",
                                         Optional savedFileList As List(Of String) = Nothing) As Boolean

            Dim arTxt As List(Of Txt) = Nothing
            If saveTexture Then
                '使用中のテクスチャを調べる
                arTxt = New List(Of Txt)
                For Each bo As BufferObject In arBO
                    If bo.TextureName = "" Then Continue For
                    Dim found As Boolean = False
                    For Each t As Txt In arTxt
                        If t.name = bo.TextureName Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found AndAlso bo._parent IsNot Nothing Then
                        Dim ti As TextureInfo = bo._parent.FindTexture(bo.TextureName)
                        If ti IsNot Nothing Then
                            arTxt.Add(New Txt(ti.name, ti.bmp, ti.name & ".bmp"))
                        Else
                            Debug.Print("BufferObjectがView3Dから切り離されているため、テクスチャを保存できません")
                        End If
                    End If
                Next
                Dim textureIndex As Integer = 1
                For Each bo As BufferObject In arBO
                    If bo.Texture Is Nothing Then Continue For
                    Dim found As Boolean = False
                    For Each t As Txt In arTxt
                        If t.bmp Is bo.Texture Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found AndAlso bo._parent IsNot Nothing Then
                        Dim ti As TextureInfo = bo._parent.FindTexture(bo.Texture)
                        If ti IsNot Nothing Then
                            Dim name As String = ti.name
                            If name = "" Then
                                name = prefix & textureIndex.ToString("D3")
                                textureIndex += 1
                            End If
                            arTxt.Add(New Txt(name, ti.bmp, name & ".bmp"))
                        Else
                            Debug.Print("BufferObjectがView3Dから切り離されているため、テクスチャを保存できません")
                        End If
                    End If
                Next
                For Each t As Txt In arTxt
                    Dim path As String = ""
                    Try
                        path = textureDirectory
                        If path.Length > 0 AndAlso path.Last <> "\"c Then path &= "\"
                        path &= t.fileName
                        Dim tmp As New Bitmap(t.bmp)
                        tmp.Save(path, Imaging.ImageFormat.Bmp)
                        If savedFileList IsNot Nothing Then savedFileList.Add(System.IO.Path.GetFullPath(path))
                    Catch ex As Exception
                        Debug.Print("テクスチャの保存に失敗しました " & path)
                        Debug.Print(ex.Message)
                    End Try
                Next
            End If

            Try
                Dim xmlWtr As New XmlTextWriter(filePath, System.Text.Encoding.UTF8)

                'XML宣言を出力
                xmlWtr.WriteStartDocument()
                xmlWtr.WriteStartElement("BufferObjects")
                xmlWtr.WriteString(Chr(13) & Chr(10))

                For Each bo As BufferObject In arBO

                    Dim nodeBO As New XmlWriteNode("BufferObject")

                    If bo.name <> "" Then nodeBO.Add("Name", bo.name)

                    If bo.Material IsNot Nothing Then
                        Dim nodeMaterial As New XmlWriteNode("Material")
                        nodeBO.Add(nodeMaterial)
                        nodeMaterial.Add("Specular", bo.Material.Specular.ToString)
                        nodeMaterial.Add("Diffuse", bo.Material.Diffuse.ToString)
                        nodeMaterial.Add("Ambient", bo.Material.Ambient.ToString)
                        nodeMaterial.Add("Shininess", bo.Material.Shininess.ToString)
                    End If

                    nodeBO.Add("Lighting", bo.Lighting)
                    nodeBO.Add("Culling", bo.Culling)
                    nodeBO.Add("Blend", bo.Blend)
                    nodeBO.Add("PointSize", bo.PointSize)
                    nodeBO.Add("PointPerspective", bo.PointPerspective)
                    nodeBO.Add("Active", bo.Active)

                    For Each m As Matrix4 In bo.Martices
                        nodeBO.Add("Martices", "{" & m.Row0.ToString & "," & m.Row1.ToString & "," & m.Row2.ToString & "," & m.Row3.ToString & "}")
                    Next

                    nodeBO.Add("_bXYZ", bo._bXYZ)
                    nodeBO.Add("_bNormal", bo._bNormal)
                    nodeBO.Add("_bTexture", bo._bTexture)
                    nodeBO.Add("_bColor", bo._bColor)
                    nodeBO.Add("_bPointSprite", bo._bPointSprite)
                    nodeBO.Add("_primitive", bo._primitive) 'Points Lines Triangles

                    Dim sb As New Text.StringBuilder
                    sb.Append("{")
                    For i As Integer = 0 To bo._ari.Count - 1
                        If i > 0 Then sb.Append(",")
                        sb.Append(bo._ari(i).ToString)
                    Next
                    sb.Append("}")
                    nodeBO.Add("_ari", sb.ToString)

                    sb.Clear()
                    sb.Append("{")
                    For i As Integer = 0 To bo._arb.Count - 1
                        If i > 0 Then sb.Append(",")
                        sb.Append("0x")
                        sb.Append(bo._arb(i).ToString("X2"))
                    Next
                    sb.Append("}")
                    nodeBO.Add("_arb", sb.ToString)
                    nodeBO.Add("_numVertex", bo._numVertex)

                    If bo.TextureName <> "" Then
                        nodeBO.Add("TextureName", bo.TextureName)
                        For Each t As Txt In arTxt
                            If t.name = bo.TextureName Then
                                nodeBO.Add("TextureFile", t.fileName)
                                Exit For
                            End If
                        Next
                    ElseIf bo.Texture IsNot Nothing Then
                        For Each t As Txt In arTxt
                            If t.bmp Is bo.Texture Then
                                nodeBO.Add("TextureName", t.name)
                                nodeBO.Add("TextureFile", t.fileName)
                                Exit For
                            End If
                        Next
                    End If

                    nodeBO.Write(xmlWtr, 1)
                Next

                xmlWtr.WriteEndElement()
                xmlWtr.WriteEndDocument()
                xmlWtr.Flush()
                xmlWtr.Close()

                If savedFileList IsNot Nothing Then savedFileList.Add(System.IO.Path.GetFullPath(filePath))
            Catch ex As Exception
                Debug.Print("BufferObjectの保存に失敗しました")
                Debug.Print(ex.Message)
                Return False
            End Try

            Return True
        End Function

        '文字列の例 {n,n,n,n,....,n}
        '"0x"形式の16進数に対応
        Private Shared Function ToIntArray(ByRef s As String) As List(Of Integer)
            Dim ar As New List(Of Integer)
            Dim i As Integer = 1
            Dim j As Integer
            While s(i) <> "}"
                If s(i) = "," Then i += 1
                j = i + 1
                While s(j) <> "," AndAlso s(j) <> "}"
                    j += 1
                End While
                Dim w As String = s.Substring(i, j - i).Trim
                If w.StartsWith("0x") Then
                    ar.Add(Convert.ToInt32(w.Substring(2), 16))
                Else
                    ar.Add(Integer.Parse(w))
                End If
                i = j
            End While
            Return ar
        End Function

        '文字列の例 {n,n,n,n,....,n}
        '"0x"形式の16進数に対応
        Private Shared Function ToByteArray(ByRef s As String) As List(Of Byte)
            Dim ar As New List(Of Byte)
            Dim i As Integer = 1
            Dim j As Integer
            While s(i) <> "}"
                If s(i) = "," Then i += 1
                j = i + 1
                While s(j) <> "," AndAlso s(j) <> "}"
                    j += 1
                End While
                Dim w As String = s.Substring(i, j - i).Trim
                If w.StartsWith("0x") Then
                    ar.Add(Convert.ToByte(w.Substring(2), 16))
                Else
                    ar.Add(Byte.Parse(w))
                End If
                i = j
            End While
            Return ar
        End Function

        ''' <summary>
        ''' SaveToXml()で保存したデータファイルを読み込みます。
        ''' loadTexture=Trueでテクスチャファイルの読み込みも行います。
        ''' 読み込まれたBufferObjectはloadedBufferObjectに、テクスチャはloadTextureListに格納されます。
        ''' これらはView3D.AddBufferObject()とBufferObject.Generate()、View3D.AddTexture()で追加しないといけません
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <param name="loadedBufferObject"></param>
        ''' <param name="loadTexture"></param>
        ''' <param name="textureDirectory"></param>
        ''' <param name="loadTextureList"></param>
        ''' <returns></returns>
        Public Shared Function LoadFromXml(filePath As String, loadedBufferObject As List(Of BufferObject),
                loadTexture As Boolean, Optional textureDirectory As String = "", Optional loadTextureList As List(Of TextureInfo) = Nothing) As Boolean

            Try
                Dim xmlDoc As New XmlDocument()
                xmlDoc.Load(filePath)

                Dim BuffferObjects As XmlNodeList = xmlDoc.GetElementsByTagName("BufferObject")
                For Each ebo As XmlElement In BuffferObjects

                    Dim name As String = ""
                    Dim n As XmlNode = ebo.SelectSingleNode("Name")
                    If n IsNot Nothing Then name = n.InnerText
                    Dim bXYZ As Boolean = Boolean.Parse(ebo.SelectSingleNode("_bXYZ").InnerText)
                    Dim bNormal As Boolean = Boolean.Parse(ebo.SelectSingleNode("_bNormal").InnerText)
                    Dim bTexture As Boolean = Boolean.Parse(ebo.SelectSingleNode("_bTexture").InnerText)
                    Dim bColor As Boolean = Boolean.Parse(ebo.SelectSingleNode("_bColor").InnerText)
                    Dim bPointSprite As Boolean = Boolean.Parse(ebo.SelectSingleNode("_bPointSprite").InnerText)
                    Dim bo As New BufferObject(Nothing, bXYZ, bNormal, bTexture, bColor, bPointSprite, name)

                    bo._primitive = Integer.Parse(ebo.SelectSingleNode("_primitive").InnerText)

                    Dim e As XmlElement = ebo.SelectSingleNode("Material")
                    If e IsNot Nothing Then
                        Dim material As New Material()
                        material.Specular = ToColor4(e.SelectSingleNode("Specular").InnerText)
                        material.Diffuse = ToColor4(e.SelectSingleNode("Diffuse").InnerText)
                        material.Ambient = ToColor4(e.SelectSingleNode("Ambient").InnerText)
                        material.Shininess = Single.Parse(e.SelectSingleNode("Shininess").InnerText)
                        bo.Material = material
                    End If

                    bo.Lighting = Boolean.Parse(ebo.SelectSingleNode("Lighting").InnerText)
                    bo.Culling = Boolean.Parse(ebo.SelectSingleNode("Culling").InnerText)
                    bo.Blend = Boolean.Parse(ebo.SelectSingleNode("Blend").InnerText)
                    bo.PointSize = Single.Parse(ebo.SelectSingleNode("PointSize").InnerText)
                    bo.PointPerspective = Boolean.Parse(ebo.SelectSingleNode("PointPerspective").InnerText)
                    bo.Active = Boolean.Parse(ebo.SelectSingleNode("Active").InnerText)

                    bo.Martices.Clear()
                    For Each matrix As XmlNode In ebo.GetElementsByTagName("Martices")
                        bo.Martices.Add(ToMatrix4(matrix.InnerText))
                    Next

                    bo._ari = ToIntArray(ebo.SelectSingleNode("_ari").InnerText)
                    bo._arb = ToByteArray(ebo.SelectSingleNode("_arb").InnerText)
                    bo._numVertex = Integer.Parse(ebo.SelectSingleNode("_numVertex").InnerText)

                    loadedBufferObject.Add(bo)

                    If loadTexture AndAlso loadTextureList IsNot Nothing Then
                        Dim ti As New TextureInfo
                        n = ebo.SelectSingleNode("TextureName")
                        If n IsNot Nothing Then
                            ti.name = n.InnerText
                            bo.TextureName = ti.name
                        End If
                        n = ebo.SelectSingleNode("TextureFile")
                        If n IsNot Nothing Then
                            Dim path As String = ""
                            Try
                                path = textureDirectory
                                If path.Length > 0 AndAlso path.Last <> "\"c Then path &= "\"
                                path &= n.InnerText
                                Dim bmp As Bitmap = Bitmap.FromFile(path)
                                ti.bmp = bmp
                                If bo.TextureName = "" Then bo.Texture = ti.bmp
                                loadTextureList.Add(ti)
                            Catch ex As Exception
                                Debug.Print("テクスチャの読み込みに失敗しました " & path)
                                Debug.Print(ex.Message)
                            End Try
                        End If
                    End If

                Next
            Catch ex As Exception
                Debug.Print("BufferObjectの読み込みに失敗しました")
                Debug.Print(ex.Message)
                Return False
            End Try

            Return True
        End Function


        ''' <summary>
        ''' 使用するテクスチャをビットマップで指定または取得します
        ''' 実際に使用するには、ビットマップはこのプロパティと、View3D.AddTexture()の両方で設定されていないといけません
        ''' </summary>
        ''' <returns></returns>
        Public Property Texture As Bitmap
            Get
                Return _bmp
            End Get
            Set(value As Bitmap)
                _bmp = value
                If value IsNot Nothing Then _textureName = "" '名前指定を無効にする
            End Set
        End Property

        ''' <summary>
        ''' 使用するテクスチャを名前で指定または取得します
        ''' 実際に使用するには、ビットマップはこのプロパティと、View3D.AddTexture()の両方で設定されていないといけません
        ''' </summary>
        ''' <returns></returns>
        Public Property TextureName As String
            Get
                Return _textureName
            End Get
            Set(value As String)
                _textureName = value
                If value <> "" Then _bmp = Nothing 'ビットマップによる指定を無効にする
            End Set
        End Property

        Public ReadOnly Property GetName As String
            Get
                Return name
            End Get
        End Property

        Friend Sub New(ByRef parent As View3D, bXYZ As Boolean, bNormal As Boolean, bTexture As Boolean, bColor As Boolean,
                       bPointSprite As Boolean, Optional name As String = "")

            If name.IndexOfAny(IO.Path.GetInvalidFileNameChars()) >= 0 Then Throw New Exception("name has invalid character")

            _parent = parent
            _bXYZ = bXYZ
            _bNormal = bNormal
            _bTexture = bTexture
            _bColor = bColor
            _bPointSprite = bPointSprite
            Me.name = name
            Martices.Add(Matrix4.Identity)
            _primitive = BeginMode.Points
        End Sub

        '頂点の大きさ(バイトサイズ)を返す
        Private Function ByteSizeOfVertex() As Integer
            Return Vector3.SizeInBytes _
                    + IIf(_bNormal, Vector3.SizeInBytes, 0) _
                    + IIf(_bTexture, Vector2.SizeInBytes, 0) _
                    + IIf(_bColor, 16, 0)
        End Function

        '_arbにSingle型データを追加する
        Private Sub arb_Add(f As Single)
            _arb.AddRange(BitConverter.GetBytes(f))
        End Sub

        '_arbにVector2型データを追加する
        Private Sub arb_Add(v As Vector2)
            _arb.AddRange(BitConverter.GetBytes(v.X))
            _arb.AddRange(BitConverter.GetBytes(v.Y))
        End Sub

        '_arbにVector3型データを追加する
        Private Sub arb_Add(v As Vector3)
            _arb.AddRange(BitConverter.GetBytes(v.X))
            _arb.AddRange(BitConverter.GetBytes(v.Y))
            _arb.AddRange(BitConverter.GetBytes(v.Z))
        End Sub

        '_arbにColor4型データを追加する
        Private Sub arb_Add(c As Color4)
            '_arb.AddRange(BitConverter.GetBytes(c.ToArgb))
            _arb.AddRange(BitConverter.GetBytes(c.R))
            _arb.AddRange(BitConverter.GetBytes(c.G))
            _arb.AddRange(BitConverter.GetBytes(c.B))
            _arb.AddRange(BitConverter.GetBytes(c.A))
        End Sub

        ''' <summary>
        ''' ポイントを追加します
        ''' ポイント以外のプリミティブが設定されていた場合はFalseを返します
        ''' </summary>
        ''' <param name="ar"></param>
        ''' <returns></returns>
        Public Function AddPoints(ar As Vector3()) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Points
                Lighting = False
            ElseIf _primitive <> BeginMode.Points Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            For Each p As Vector3 In ar
                '頂点バッファに追加
                arb_Add(p)
                If _bNormal Then arb_Add(Vector3.Zero)
                If _bTexture Then arb_Add(Vector2.Zero)
                If _bColor Then arb_Add(Color4.Black)

                'インデックスを追加
                _ari.Add(_numVertex)
                _numVertex += 1
            Next
            Return True
        End Function

        ''' <summary>
        ''' ポイントと頂点カラーを追加します
        ''' ポイント以外のプリミティブが設定されていた場合はFalseを返します
        ''' c1: 頂点カラー（頂点バッファが頂点カラーを持っていない場合は設定されません）
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <param name="c1"></param>
        ''' <returns></returns>
        Public Function AddPoint(p1 As Vector3, c1 As Color4) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Points
                Lighting = False
            ElseIf _primitive <> BeginMode.Points Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            '頂点バッファに追加
            arb_Add(p1)
            If _bNormal Then arb_Add(Vector3.Zero)
            If _bTexture Then arb_Add(Vector2.Zero)
            If _bColor Then arb_Add(c1)

            'インデックスを追加
            _ari.Add(_numVertex)
            _numVertex += 1
            Return True
        End Function

        ''' <summary>
        ''' ラインと頂点カラーを追加します
        ''' ライン以外のプリミティブが設定されていた場合はFalseを返します
        ''' c1,c2: p1,p2に対応する頂点カラー（頂点バッファが頂点カラーを持っていない場合は設定されません）
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <param name="c1"></param>
        ''' <param name="p2"></param>
        ''' <param name="c2"></param>
        ''' <returns></returns>
        Public Function AddLine(p1 As Vector3, c1 As Color4, p2 As Vector3, c2 As Color4) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Lines
                Lighting = False
            ElseIf _primitive <> BeginMode.Lines Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            '頂点バッファに追加
            arb_Add(p1)
            If _bNormal Then arb_Add(Vector3.Zero)
            If _bTexture Then arb_Add(Vector2.Zero)
            If _bColor Then arb_Add(c1)
            arb_Add(p2)
            If _bNormal Then arb_Add(Vector3.Zero)
            If _bTexture Then arb_Add(Vector2.Zero)
            If _bColor Then arb_Add(c2)

            'インデックスを追加
            _ari.Add(_numVertex)
            _ari.Add(_numVertex + 1)
            _numVertex += 2
            Return True
        End Function

        ''' <summary>
        ''' 三角形と頂点カラーを追加します
        ''' 三角形以外のプリミティブが設定されていた場合はFalseを返します
        ''' c1,c2,c3: p1,p2,p3に対応する頂点カラー（頂点バッファが頂点カラーを持っていない場合は設定されません）
        ''' 頂点バッファが法線ベクトルを持っている場合は、各頂点からの計算値が代入されます。
        ''' 三角形は時計回りの頂点の並びが「表面」として処理されます
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <param name="c1"></param>
        ''' <param name="p2"></param>
        ''' <param name="c2"></param>
        ''' <param name="p3"></param>
        ''' <param name="c3"></param>
        ''' <returns></returns>
        Public Function AddTriangle(p1 As Vector3, c1 As Color4, p2 As Vector3, c2 As Color4, p3 As Vector3, c3 As Color4) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Triangles
                Lighting = True
            ElseIf _primitive <> BeginMode.Triangles Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            '必要に応じて法線を計算
            Dim vn As Vector3
            If _bNormal Then
                vn = Vector3.Cross(p3 - p1, p2 - p3)
                vn.Normalize()
            End If

            '頂点バッファに追加
            arb_Add(p1)
            If _bNormal Then arb_Add(vn)
            If _bTexture Then arb_Add(Vector2.Zero)
            If _bColor Then arb_Add(c1)
            arb_Add(p2)
            If _bNormal Then arb_Add(vn)
            If _bTexture Then arb_Add(Vector2.Zero)
            If _bColor Then arb_Add(c2)
            arb_Add(p3)
            If _bNormal Then arb_Add(vn)
            If _bTexture Then arb_Add(Vector2.Zero)
            If _bColor Then arb_Add(c3)

            'インデックスを追加
            _ari.Add(_numVertex)
            _ari.Add(_numVertex + 1)
            _ari.Add(_numVertex + 2)
            _numVertex += 3

            Return True
        End Function

        ''' <summary>
        ''' 三角形とテクスチャ座標を追加します
        ''' 三角形以外のプリミティブが設定されていた場合はFalseを返します
        ''' t1,t2,t3: p1,p2,p3に対応するテクスチャ座標（頂点バッファがテクスチャ座標を持っていない場合は設定されません）
        ''' 頂点バッファが法線ベクトルを持っている場合は、各頂点からの計算値が代入されます。
        ''' 三角形は時計回りの頂点の並びが「表面」として処理されます
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <param name="t1"></param>
        ''' <param name="p2"></param>
        ''' <param name="t2"></param>
        ''' <param name="p3"></param>
        ''' <param name="t3"></param>
        ''' <returns></returns>
        Public Function AddTriangle(p1 As Vector3, t1 As Vector2, p2 As Vector3, t2 As Vector2, p3 As Vector3, t3 As Vector2) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Triangles
                Lighting = True
            ElseIf _primitive <> BeginMode.Triangles Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            '必要に応じて法線を計算
            Dim vn As Vector3
            If _bNormal Then
                vn = Vector3.Cross(p3 - p1, p2 - p3)
                vn.Normalize()
            End If

            '頂点バッファに追加
            arb_Add(p1)
            If _bNormal Then arb_Add(vn)
            If _bTexture Then arb_Add(t1)
            If _bColor Then arb_Add(Color4.White)
            arb_Add(p2)
            If _bNormal Then arb_Add(vn)
            If _bTexture Then arb_Add(t2)
            If _bColor Then arb_Add(Color4.White)
            arb_Add(p3)
            If _bNormal Then arb_Add(vn)
            If _bTexture Then arb_Add(t3)
            If _bColor Then arb_Add(Color4.White)

            'インデックスを追加
            _ari.Add(_numVertex)
            _ari.Add(_numVertex + 1)
            _ari.Add(_numVertex + 2)
            _numVertex += 3
            Return True
        End Function

        ''' <summary>
        ''' 四角形とテクスチャ座標を追加します
        ''' １つの四角形は２つの三角形で表されます
        ''' 三角形以外のプリミティブが設定されていた場合はFalseを返します
        ''' c1,c2,c3,c4: p1～p4に対応する頂点カラー（頂点バッファが頂点カラーを持っていない場合は設定されません）
        ''' div: 四角形を分割して作成します。例えば、div=2の場合、4毎の四角形⇒8毎の三角形が作成されます。デフォルトは1で分割されません
        ''' 頂点バッファが法線ベクトルを持っている場合は、各頂点からの計算値が代入されます。
        ''' 四角形は時計回りの頂点の並びが「表面」として処理されます
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <param name="c1"></param>
        ''' <param name="p2"></param>
        ''' <param name="c2"></param>
        ''' <param name="p3"></param>
        ''' <param name="c3"></param>
        ''' <param name="p4"></param>
        ''' <param name="c4"></param>
        ''' <param name="div"></param>
        ''' <returns></returns>
        Public Function AddQuad(p1 As Vector3, c1 As Color4, p2 As Vector3, c2 As Color4, p3 As Vector3, c3 As Color4,
                                p4 As Vector3, c4 As Color4, Optional div As Integer = 1) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Triangles
                Lighting = False
            ElseIf _primitive <> BeginMode.Triangles Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            '必要に応じて法線を計算
            Dim vn As Vector3
            If _bNormal Then
                vn = Vector3.Cross(p3 - p1, p2 - p3)
                vn.Normalize()
            End If

            If div < 1 Then div = 1
            For j As Integer = 0 To div
                Dim vL As New Vector3(
                    (p4.X - p1.X) * j / div + p1.X,
                    (p4.Y - p1.Y) * j / div + p1.Y,
                    (p4.Z - p1.Z) * j / div + p1.Z)
                Dim vR As New Vector3(
                    (p3.X - p2.X) * j / div + p2.X,
                    (p3.Y - p2.Y) * j / div + p2.Y,
                    (p3.Z - p2.Z) * j / div + p2.Z)
                Dim cL As New Color4(
                    (c4.R - c1.R) * j / div + c1.R,
                    (c4.G - c1.G) * j / div + c1.G,
                    (c4.B - c1.B) * j / div + c1.B,
                    (c4.A - c1.A) * j / div + c1.A)
                Dim cR As New Color4(
                    (c3.R - c2.R) * j / div + c2.R,
                    (c3.G - c2.G) * j / div + c2.G,
                    (c3.B - c2.B) * j / div + c2.B,
                    (c3.A - c2.A) * j / div + c2.A)
                For i As Integer = 0 To div
                    Dim p As New Vector3(
                        (vR.X - vL.X) * i / div + vL.X,
                        (vR.Y - vL.Y) * i / div + vL.Y,
                        (vR.Z - vL.Z) * i / div + vL.Z)
                    Dim c As New Color4(
                        (cR.R - cL.R) * i / div + cL.R,
                        (cR.G - cL.G) * i / div + cL.G,
                        (cR.B - cL.B) * i / div + cL.B,
                        (cR.A - cL.A) * i / div + cL.A)

                    '頂点バッファに追加
                    arb_Add(p)
                    If _bNormal Then arb_Add(vn)
                    If _bTexture Then arb_Add(New Vector2(0, 0))
                    If _bColor Then arb_Add(c)
                Next
            Next

            For j As Integer = 0 To div - 1
                For i As Integer = 0 To div - 1

                    'インデックスを追加
                    _ari.Add(_numVertex + j * (div + 1) + i)            '左上
                    _ari.Add(_numVertex + j * (div + 1) + i + 1)        '右上
                    _ari.Add(_numVertex + (j + 1) * (div + 1) + i + 1)  '右下
                    _ari.Add(_numVertex + j * (div + 1) + i)            '左上
                    _ari.Add(_numVertex + (j + 1) * (div + 1) + i + 1)  '右下
                    _ari.Add(_numVertex + (j + 1) * (div + 1) + i)      '左下
                Next
            Next
            _numVertex += (div + 1) * (div + 1)
            Return True
        End Function

        ''' <summary>
        ''' 四角形とテクスチャ座標を追加します
        ''' １つの四角形は２つの三角形で表されます
        ''' 三角形以外のプリミティブが設定されていた場合はFalseを返します
        ''' t1,t2,t3,t4: p1～p4に対応するテクスチャ座標（頂点バッファがテクスチャ座標を持っていない場合は設定されません）
        ''' div: 四角形を分割して作成します。例えば、div=2の場合、4毎の四角形⇒8毎の三角形が作成されます。デフォルトは1で分割されません
        ''' 頂点バッファが法線ベクトルを持っている場合は、各頂点からの計算値が代入されます。
        ''' 四角形は時計回りの頂点の並びが「表面」として処理されます
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <param name="t1"></param>
        ''' <param name="p2"></param>
        ''' <param name="t2"></param>
        ''' <param name="p3"></param>
        ''' <param name="t3"></param>
        ''' <param name="p4"></param>
        ''' <param name="t4"></param>
        ''' <param name="div"></param>
        ''' <returns></returns>
        Public Function AddQuad(p1 As Vector3, t1 As Vector2, p2 As Vector3, t2 As Vector2, p3 As Vector3, t3 As Vector2,
                                p4 As Vector3, t4 As Vector2, Optional div As Integer = 1) As Boolean
            If _numVertex = 0 Then
                _primitive = BeginMode.Triangles
                Lighting = True
            ElseIf _primitive <> BeginMode.Triangles Then
                Debug.Print("ポイント、ライン、トライアングルなどのプリミティブを１つのBufferObjectに混在させることはできません")
                Return False
            End If

            '必要に応じて法線を計算
            Dim vn As Vector3
            If _bNormal Then
                vn = Vector3.Cross(p3 - p1, p2 - p3)
                vn.Normalize()
            End If

            If div < 1 Then div = 1
            For j As Integer = 0 To div
                Dim vL As New Vector3(
                    (p4.X - p1.X) * j / div + p1.X,
                    (p4.Y - p1.Y) * j / div + p1.Y,
                    (p4.Z - p1.Z) * j / div + p1.Z)
                Dim vR As New Vector3(
                    (p3.X - p2.X) * j / div + p2.X,
                    (p3.Y - p2.Y) * j / div + p2.Y,
                    (p3.Z - p2.Z) * j / div + p2.Z)
                Dim tL As New Vector2(
                    (t4.X - t1.X) * j / div + t1.X,
                    (t4.Y - t1.Y) * j / div + t1.Y)
                Dim tR As New Vector2(
                    (t3.X - t2.X) * j / div + t2.X,
                    (t3.Y - t2.Y) * j / div + t2.Y)
                For i As Integer = 0 To div
                    Dim p As New Vector3(
                        (vR.X - vL.X) * i / div + vL.X,
                        (vR.Y - vL.Y) * i / div + vL.Y,
                        (vR.Z - vL.Z) * i / div + vL.Z)
                    Dim t As New Vector2(
                        (tR.X - tL.X) * i / div + tL.X,
                        (tR.Y - tL.Y) * i / div + tL.Y)

                    '頂点バッファに追加
                    arb_Add(p)
                    If _bNormal Then arb_Add(vn)
                    If _bTexture Then arb_Add(t)
                    If _bColor Then arb_Add(Color4.White)
                Next
            Next

            For j As Integer = 0 To div - 1
                For i As Integer = 0 To div - 1

                    'インデックスを追加
                    _ari.Add(_numVertex + j * (div + 1) + i)            '左上
                    _ari.Add(_numVertex + j * (div + 1) + i + 1)        '右上
                    _ari.Add(_numVertex + (j + 1) * (div + 1) + i + 1)  '右下
                    _ari.Add(_numVertex + j * (div + 1) + i)            '左上
                    _ari.Add(_numVertex + (j + 1) * (div + 1) + i + 1)  '右下
                    _ari.Add(_numVertex + (j + 1) * (div + 1) + i)      '左下
                Next
            Next
            _numVertex += (div + 1) * (div + 1)

            '頂点バッファに追加
            'arb_Add(p1)
            'If _bNormal Then arb_Add(vn)
            'If _bTexture Then arb_Add(t1)
            'If _bColor Then arb_Add(Color4.White)
            'arb_Add(p2)
            'If _bNormal Then arb_Add(vn)
            'If _bTexture Then arb_Add(t2)
            'If _bColor Then arb_Add(Color4.White)
            'arb_Add(p3)
            'If _bNormal Then arb_Add(vn)
            'If _bTexture Then arb_Add(t3)
            'If _bColor Then arb_Add(Color4.White)
            'arb_Add(p4)
            'If _bNormal Then arb_Add(vn)
            'If _bTexture Then arb_Add(t4)
            'If _bColor Then arb_Add(Color4.White)

            'インデックスを追加
            '_ari.Add(_numVertex)
            '_ari.Add(_numVertex + 1)
            '_ari.Add(_numVertex + 2)
            '_ari.Add(_numVertex)
            '_ari.Add(_numVertex + 2)
            '_ari.Add(_numVertex + 3)
            '_numVertex += 4
            Return True
        End Function

        ''' <summary>
        ''' AddLine()などで追加された頂点データをOpenGLのバッファに設定します。
        ''' 既にバッファに設定している場合は、一旦削除して、最新の頂点データで再設定を行います。
        ''' これによりBufferObjectが表示されるようになります。
        ''' </summary>
        Public Sub Generate()

            '頂点バッファが登録済みであれば削除する
            If _vao <> 0 Then
                GL.DeleteVertexArray(_vao)
                _vao = 0
            End If
            If _vbo <> 0 Then
                GL.DeleteBuffers(1, _vbo)
                _vbo = 0
            End If
            If _ibo <> 0 Then
                GL.DeleteBuffers(1, _ibo)
                _ibo = 0
                _numIndex = 0
            End If

            If _arb.Count > 0 Then
                'VBOを1コ作成し、頂点データを送り込む
                GL.GenBuffers(1, _vbo)
                GL.BindBuffer(BufferTarget.ArrayBuffer, _vbo)
                Dim data As IntPtr = Runtime.InteropServices.Marshal.AllocCoTaskMem(_arb.Count)
                Runtime.InteropServices.Marshal.Copy(_arb.ToArray, 0, data, _arb.Count)
                GL.BufferData(BufferTarget.ArrayBuffer, New IntPtr(_arb.Count), data, BufferUsageHint.StaticDraw)
                Runtime.InteropServices.Marshal.FreeCoTaskMem(data)

                '描画時に必要なパラメータをコピー
                _numIndex = _ari.Count
                _bm = _primitive

                'IBOを1コ作成し、インデックスデータを送り込む
                GL.GenBuffers(1, _ibo)
                GL.BindBuffer(BufferTarget.ElementArrayBuffer, _ibo)
                data = Runtime.InteropServices.Marshal.AllocCoTaskMem(_numIndex << 2)
                Runtime.InteropServices.Marshal.Copy(_ari.ToArray, 0, data, _numIndex)
                GL.BufferData(BufferTarget.ElementArrayBuffer, New IntPtr(_numIndex << 2), data, BufferUsageHint.StaticDraw)
                Runtime.InteropServices.Marshal.FreeCoTaskMem(data)

                'VAOを1コ作成し、設定
                GL.GenVertexArrays(1, _vao)
                GL.BindVertexArray(_vao)

                ''各Arrayを有効化
                'GL.EnableClientState(ArrayCap.VertexArray)
                'If _bNormal Then
                '    GL.EnableClientState(ArrayCap.NormalArray)
                'End If
                'If _bTexture Then
                '    GL.EnableClientState(ArrayCap.TextureCoordArray)
                'End If
                'If _bColor Then
                '    GL.EnableClientState(ArrayCap.ColorArray)
                'End If

                GL.BindBuffer(BufferTarget.ArrayBuffer, _vbo)

                '頂点の位置、法線、テクスチャ情報の場所を指定
                GL.VertexPointer(3, VertexPointerType.Float, ByteSizeOfVertex, 0)
                Dim offset As Integer = Vector3.SizeInBytes
                If _bNormal Then
                    GL.NormalPointer(NormalPointerType.Float, ByteSizeOfVertex, offset)
                    offset += Vector3.SizeInBytes
                End If
                If _bTexture Then
                    GL.TexCoordPointer(2, TexCoordPointerType.Float, ByteSizeOfVertex, offset)
                    offset += Vector2.SizeInBytes
                End If
                If _bColor Then
                    GL.ColorPointer(4, ColorPointerType.Float, ByteSizeOfVertex, offset)
                End If

                ''各Arrayを無効化
                'GL.DisableClientState(ArrayCap.VertexArray)
                'If _bNormal Then
                '    GL.DisableClientState(ArrayCap.NormalArray)
                'End If
                'If _bTexture Then
                '    GL.DisableClientState(ArrayCap.TextureCoordArray)
                'End If
                'If _bColor Then
                '    GL.DisableClientState(ArrayCap.ColorArray)
                'End If

                GL.BindBuffer(BufferTarget.ArrayBuffer, 0)
                GL.BindBuffer(BufferTarget.ElementArrayBuffer, 0)
                GL.BindVertexArray(0)
            End If

            'If bTexture Then
            '    'テクスチャが登録済みであれば削除する
            '    If _txt <> 0 Then
            '        GL.DeleteTexture(_txt)
            '        _txt = 0
            '    End If

            '    If _bmp IsNot Nothing Then

            '        'テクスチャ
            '        GL.Enable(EnableCap.Texture2D)
            '        _txt = GL.GenTexture()
            '        GL.BindTexture(TextureTarget.Texture2D, _txt)
            '        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMinFilter, TextureMinFilter.Nearest)
            '        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMagFilter, TextureMagFilter.Nearest)
            '        Dim bmpData As Imaging.BitmapData = _bmp.LockBits(New Rectangle(0, 0, _bmp.Width, _bmp.Height), Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            '        GL.TexImage2D(TextureTarget.Texture2D, 0, PixelInternalFormat.Rgba, bmpData.Width, bmpData.Height, 0, OpenTK.Graphics.OpenGL.PixelFormat.Bgra, PixelType.UnsignedByte, bmpData.Scan0)
            '        _bmp.UnlockBits(bmpData)
            '        GL.BindTexture(TextureTarget.Texture2D, 0)
            '    End If
            'End If
        End Sub

    End Class

End Class
