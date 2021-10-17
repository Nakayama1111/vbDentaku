Public Class FomDentaku

    ''' <summary>
    ''' 画面に数値として表示する文字列
    ''' </summary>
    Dim _dispNumString As String = 0
    ''' <summary>
    ''' 画面に表示する文字列を数値として保存(メソッドで直接変更はしない)
    ''' </summary>
    Dim dispNum As Decimal
    ''' <summary>
    ''' 1つ目の数値を保持
    ''' </summary>
    Dim fromNum As Decimal = 0
    ''' <summary>
    ''' 2つ目の数値を保持
    ''' </summary>
    Dim toNum As Decimal = 0
    ''' <summary>
    ''' 直前に入力された情報の種類を記録する(初期値はnullとする)
    ''' </summary>
    Dim beforInpType As InputType = InputType.null
    ''' <summary>
    ''' 入力中の演算子を記憶する(初期値はcalcNullとする)
    ''' </summary>
    Dim beforCalc As CalcType = CalcType.calcNull
    ''' <summary>
    ''' <summary>
    ''' 画面に表示する数値のリセットフラグ trueであれば、数値入力時に表示数値をリセットする(初期値はfalseとする)
    ''' </summary>
    Dim dispNumResetFlug As Boolean = False
    ''' <summary>
    ''' 2つ目の数値更新フラグ trueであれば、入力中の数値を2つ目の数値とする(初期値はfalseとする)
    ''' </summary>
    Dim toNumUpdateFlug As Boolean = False
    ''' <summary>
    ''' エラーメッセージ表示フラグ true = エラーメッセージ表示中 false = エラー無し(初期値をfalseとする)
    ''' </summary>
    Dim errorFlug As Boolean = False

    ''' <summary>
    ''' 画面に表示する文字と計算用の数値を同一にするプロパティ
    ''' </summary>
    ''' <returns></returns>
    Public Property dispNumString As String
        Get
            Return _dispNumString
        End Get

        Set(value As String)

            'valueを数値に変換できない場合、文字列をエラーメッセージとしてエラー処理を行う
            If IsNumeric(value) = False Then
                LblDisp.Text = value

                errorFlug = True
                Exit Property
            End If

            _dispNumString = value
            LblDisp.Text = value

            '最後の文字が[.]の場合、整数として保存
            If Strings.Right(value, 1).Equals(".") Then
                value = value.Substring(0, value.IndexOf("."))
            End If
            dispNum = CType(value, Decimal)

        End Set
    End Property

    ''' <summary>
    ''' 演算記号
    ''' </summary>
    Public Enum CalcType
        ''' <summary>
        ''' 和算演算子
        ''' </summary>
        plus
        ''' <summary>
        ''' 減算演算子
        ''' </summary>
        minus
        ''' <summary>
        ''' 乗算演算子
        ''' </summary>
        multiplied
        ''' <summary>
        ''' 除算演算子
        ''' </summary>
        divided
        ''' <summary>
        ''' 初期値
        ''' </summary>
        calcNull
    End Enum

    ''' <summary>
    ''' 入力された情報の種類
    ''' </summary>
    Public Enum InputType
        ''' <summary>
        ''' 数値
        ''' </summary>
        num
        ''' <summary>
        ''' 演算子
        ''' </summary>
        calc
        ''' <summary>
        ''' イコール
        ''' </summary>
        equale
        ''' <summary>
        ''' 無し
        ''' </summary>
        null
    End Enum

    ''' <summary>
    ''' CalcTypeを記号に変換するテーブル
    ''' </summary>
    Private CalcTypeSymple As New Dictionary(Of CalcType, String) From
        {
    {CalcType.plus, "+"},
    {CalcType.minus, "-"},
    {CalcType.multiplied, "×"},
    {CalcType.divided, "÷"}
    }

    ''' <summary>
    ''' ロード時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dispNumString = 0
        lblFormula.Text = ""
    End Sub

    ''' <summary>
    ''' 右端の数字を削除
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnBac_Click(sender As Object, e As EventArgs) Handles BtnBac.Click
        'エラーの場合、または式表示用ラベルに[=]が表示されている場合はBtnCの処理を行う
        If errorFlug = True OrElse lblFormula.Text.Contains("=") Then
            BtnC.PerformClick()
        ElseIf dispNumString.Count = 1 Then
            '桁数が1桁の場合、表示数値を0にする
            dispNumString = 0
        Else
            dispNumString = dispNumString.Substring(0, dispNumString.Count - 1)
        End If
    End Sub

    ''' <summary>
    ''' 入力中の数値のみをクリア
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnCE_Click(sender As Object, e As EventArgs) Handles BtnCE.Click
        'エラーの場合、または式表示用ラベルに[=]が表示されている場合はBtnCの処理を行う
        If errorFlug = True OrElse lblFormula.Text.Contains("=") Then
            BtnC.PerformClick()
        Else
            dispNumString = 0
        End If
    End Sub

    ''' <summary>
    ''' 初期状態に戻す
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnC_Click(sender As Object, e As EventArgs) Handles BtnC.Click
        fromNum = 0
        toNum = 0
        dispNumString = 0
        lblFormula.Text = ""
        beforInpType = InputType.null
        beforCalc = CalcType.calcNull
        dispNumResetFlug = False
        toNumUpdateFlug = False
        errorFlug = False
    End Sub


    ''' <summary>
    ''' 数字ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Num_Click(sender As Object, e As EventArgs) Handles BtnNum0.Click, BtnNum1.Click,
        BtnNum2.Click, BtnNum3.Click, BtnNum4.Click, BtnNum5.Click, BtnNum6.Click, BtnNum7.Click,
        BtnNum8.Click, BtnNum9.Click
        Num_Process(sender.text)
    End Sub

    ''' <summary>
    ''' dot(小数点)ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnDot_Click(sender As Object, e As EventArgs) Handles BtnDot.Click
        'エラーメッセージ表示中は処理しない
        If errorFlug = True Then
            Exit Sub
        End If

        '画面表示数値のリセットフラグがtrueであれば、表示数値を[0.]とする
        If dispNumResetFlug = True Then
            dispNumString = "0."
            dispNumResetFlug = False
        Else
            '画面表示数値のリセットフラグがFalseであれば、表示数値に小数点が無い場合のみ小数点を追加する
            If dispNumString.Contains(".") = False Then
                dispNumString = dispNumString & "."
            End If

        End If

        '前回の入力を数値にする
        beforInpType = InputType.num
    End Sub

    ''' <summary>
    ''' ±ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnNumMineus_Click(sender As Object, e As EventArgs) Handles BtnNumMineus.Click
        dispNumString = -dispNum
    End Sub

    ''' <summary>
    ''' 和算ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnPlus_Click(sender As Object, e As EventArgs) Handles BtnPlus.Click
        Calc_Get(CalcType.plus)
    End Sub

    ''' <summary>
    ''' 減算ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnMineus_Click(sender As Object, e As EventArgs) Handles BtnMineus.Click
        Calc_Get(CalcType.minus)
    End Sub

    ''' <summary>
    ''' 乗算ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnMulti_Click(sender As Object, e As EventArgs) Handles BtnMulti.Click
        Calc_Get(CalcType.multiplied)
    End Sub

    ''' <summary>
    ''' 除算ボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnDivide_Click(sender As Object, e As EventArgs) Handles BtnDivide.Click
        Calc_Get(CalcType.divided)
    End Sub

    ''' <summary>
    ''' イコールボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnEqule_Click(sender As Object, e As EventArgs) Handles BtnEqule.Click

        'エラーメッセージ表示中は処理しない
        If errorFlug = True Then
            Exit Sub
        End If

        '前回の演算子が存在するか確認する
        If beforCalc = CalcType.calcNull Then
            '前回の演算子が存在しない場合
            '式表示用ラベルの文字列を[数値 =]に変更する
            lblFormula.Text = dispNum & " ="
        Else
            '前回の演算子が存在する場合
            '2つ目の数値入力フラグがtrueの場合、画面に表示中の数値を2つ目の数値とする
            If toNumUpdateFlug = True Then
                toNum = dispNum
                toNumUpdateFlug = False
            Else
                '2つ目の数値入力フラグがfalseの場合、画面に表示中の数値を1つ目の数値とする
                fromNum = dispNum
            End If
            '前回の入力演算子に応じた計算を実行する

            dispNumString = Calc_Process(fromNum, toNum, beforCalc)

            '式表示用ラベルの文字列を[1つ目の数値 演算子 2つ目の数値 = 演算結果]とする
            lblFormula.Text = fromNum & " " & CalcTypeSymple(beforCalc) & " " & toNum & " = " & dispNum

            '現在の式を履歴に追加する
            dgvRireki.Rows.Add(lblFormula.Text)
        End If

        '直前の入力をイコールとする
        beforInpType = InputType.equale
        '表示数値リセットフラグをtrueにする
        dispNumResetFlug = True
        'エラーが発生している場合、式表示用ラベルをブランクにする
        If errorFlug = True Then
            lblFormula.Text = ""
        End If
    End Sub

    ''' <summary>
    ''' 数値の入力を受け取り、表示数値を変更する
    ''' </summary>
    ''' <param name="impNum">入力数値</param>
    Private Sub Num_Process(impNum As Integer)
        'エラーメッセージ表示中は処理しない
        If errorFlug = True Then
            Exit Sub
        End If

        '直前の入力を数値とする
        beforInpType = InputType.num

        '式表示用ラベルに[=]があれば、その文字列をブランクにする
        If 0 <= lblFormula.Text.IndexOf("=") Then
            lblFormula.Text = ""
        End If

        '表紙数値リセットフラグがtrueの場合、又は表示数値が[0]の場合表示する数値を入力数値で初期化する
        If dispNumResetFlug = True OrElse dispNumString.Equals("0") Then
            dispNumString = impNum
            dispNumResetFlug = False
            Exit Sub
        End If

        '入力できる数値を15桁までとする処理
        If dispNum.ToString.Count >= 15 Then
            dispNumString = dispNum.ToString.Substring(0, dispNum.ToString.Count - 1) & impNum
        Else
            'lblDispへ表示する数値の右に入力数値を追加する
            dispNumString = dispNumString & impNum
        End If
    End Sub

    ''' <summary>
    ''' 四則演算の演算子入力を受け取り、前回の演算子があれば計算処理を行うメソッドを起動する
    ''' </summary>
    ''' <param name="calc">入力演算子</param>
    Private Sub Calc_Get(calc As CalcType)

        'エラーメッセージ表示中は処理しない
        If errorFlug = True Then
            Exit Sub
        End If

        '表示数値が[n.]となっていた場合、表示数値を[n]にする処理
        LblDisp.Text = dispNum


        '演算子を入力済みかつ、直前の入力が演算子やイコールでない場合かつ、
        '式表示用ラベルがブランクではない場合、前回の入力演算子に応じた計算を実行する
        If beforCalc <> CalcType.calcNull AndAlso beforInpType <> InputType.calc AndAlso
        beforInpType <> InputType.equale AndAlso lblFormula.Text.Equals("") = False Then
            dispNumString = Calc_Process(fromNum, dispNum, beforCalc)
        End If

        '入力演算子と数値を記憶する
        beforCalc = calc
        fromNum = dispNum

        '式表示用ラベルの文字列を[数値 演算子]に変更する
        lblFormula.Text = fromNum & " " & CalcTypeSymple(calc)

        '直前の入力を演算子とし、表示数値リセットフラグをtrueとする
        beforInpType = InputType.calc
        dispNumResetFlug = True

        '2つ目の数値入力フラグをtrueにする
        toNumUpdateFlug = True

        'エラーが発生している場合、式表示用ラベルをブランクにする
        If errorFlug = True Then
            lblFormula.Text = ""
        End If
    End Sub


    ''' <summary>
    ''' 演算処理を行う
    ''' </summary>
    ''' <param name="calc">演算子</param>
    Private Function Calc_Process(locFromNum As Decimal, locToNum As Decimal, calc As CalcType) As String
        Dim answer As Decimal
        Try
            Select Case beforCalc
            '演算子が+の計算
                Case CalcType.plus
                    answer = locFromNum + locToNum
            '演算子が-の計算
                Case CalcType.minus
                    answer = locFromNum - locToNum
            '演算子が×の計算
                Case CalcType.multiplied
                    answer = locFromNum * locToNum
            '演算子が÷の計算
                Case CalcType.divided
                    If locToNum = 0 Then
                        Return "0で割ることは出来ません"
                    End If
                    answer = locFromNum / locToNum
                    '答えが15桁を超える場合は切り落とす
                    If answer.ToString.Count > 15 Then
                        answer = Math.Round(answer, 15 - answer.ToString.IndexOf(".") - 1)
                    End If
                Case Else
                    Return "Calc_Processの入力演算子が不正です"
            End Select
        Catch ex As OverflowException
            Return "オーバーフローしました"
        End Try

        '0の少数がある場合、それを削除して返す
        If answer Mod 1 = 0 AndAlso answer.ToString.Contains(".") Then
                Return answer.ToString.Substring(0, answer.ToString.IndexOf("."))
            End If

            '桁数が15を超える場合はエラーとする
            If answer.ToString.Count > 15 Then
                Return "オーバーフローしました"
            End If

            Return answer

    End Function

    ''' <summary>
    ''' キーの入力を受け取る
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TabControl1_KeyDown(sender As Object, e As KeyEventArgs) Handles TabControl.KeyDown
        '押されたキーの確認
        Select Case e.KeyCode
            '0～9のキーが押された場合
            Case Keys.D0 To Keys.D9
                If TabControl.SelectedTab.Name.Equals("DentakuPage") Then
                    Dim num As Char = Chr(e.KeyCode)
                    Num_Process(Convert.ToInt32(num.ToString))
                End If
                '=が入力された場合
            Case Keys.OemMinus
                If e.Modifiers = Keys.Shift Then
                    BtnEqule.PerformClick()
                Else
                    '-が押された場合
                    BtnMineus.PerformClick()
                End If
                '+が入力された場合
            Case Keys.Oemplus
                If e.Modifiers = Keys.Shift Then
                    BtnPlus.PerformClick()
                End If
                '*が入力された場合
            Case Keys.Oem1
                If e.Modifiers = Keys.Shift Then
                    BtnMulti.PerformClick()
                End If
                '/が入力された場合
            Case Keys.Oem2
                BtnDivide.PerformClick()
                'Enterが入力された場合
            Case Keys.Enter
                BtnEqule.PerformClick()
                '.が入力された場合
            Case Keys.OemPeriod
                BtnDot.PerformClick()
                'BackSpaceが入力された場合
            Case Keys.Back
                BtnBac.PerformClick()
                'Deletが入力された場合
            Case Keys.Delete
                BtnCE.PerformClick()
                'Escapeが入力された場合
            Case Keys.Escape
                BtnC.PerformClick()
        End Select
    End Sub

End Class
