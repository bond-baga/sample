Option Explicit On


Public Class DataGridViewEx
  Inherits DataGridView

  ' ① DataGridViewにセルがコンボボックスの場合に発生するイベントを追加する
  Public Event SelectedIndexChanged(ByVal sender As Object, ByVal e As ExDataGridViewCellEventArgs)

  ' ② RowValidatingイベントはタブの切り替えなどでも発生するので
  ' その行内のセルが編集モードになった場合とその行でキー入力が行われた場合のみ
  ' RowValidatingイベントが発生するようにする。
  Private mChanged As Boolean = False

  ' ③ セルのEnter時の値を保持
  ' アプリ側ではCurrentValueプロパティで値を取得する
  Private mCurrentValue As String

  ' ④ ユーザ定義のイベントCellMove時（KeyDown時に発生させる）に編集中のセルの値を取得する必要がある
  ' ０：編集なし　１：テキストボックスの編集　２：コンボボックスの編集
  ' 
  'Private mEditingState As Integer = 0

  ' ⑤ セルのEnter時の値を保持
  ' アプリ側ではCurrentValueプロパティで値を取得する
  Public Event CellMove(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

  ' ⑥ 
  ' アプリ側でRowRemovingイベントを発生させる。キャンセルが可能。-> UserDeletingRowイベントが備わっている
  'Public Event RowRemoving(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs)

  Private mItem(20) As Object

  ' ボタンに移動できるどうか
  Private mIsButtonViable As Boolean
  ' 

  Private Enum Tsugi As Integer
    NASHI = 0
    MAE = 1
    USHIRO = 2
    MIGI = 3
    HIDARI = 4
    UE = 5
    SHITA = 6
  End Enum

  'Protected Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
  '  ' 最初または最後のセルでコントロールを抜ける場合にこのプロシージャがコールされる

  '  'If (keyData And Keys.KeyCode) = Keys.Tab Then
  '  '  If (keyData And Keys.Modifiers) = Keys.Shift Then
  '  '  Else
  '  '  End If
  '  'End If
  '  Return MyBase.ProcessDialogKey(keyData)
  'End Function

  'Protected Overrides Function IsInputKey(ByVal keyData As Keys) As Boolean
  '  Return MyBase.IsInputKey(keyData)
  'End Function

  Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
    Dim NextStep As Tsugi

    NextStep = Tsugi.NASHI

    If Me Is Nothing OrElse Me.ReadOnly = True OrElse Me.CurrentCell Is Nothing Then
      Return MyBase.ProcessCmdKey(msg, keyData)
    End If

    If (keyData And Keys.KeyCode) = Keys.Enter And (keyData And Keys.Modifiers) = Keys.Shift Then
      ' セル内で改行する場合はShift + Enterなので処理をしない
      Return MyBase.ProcessCmdKey(msg, keyData)
    End If

    If (keyData And Keys.KeyCode) = Keys.Tab Then
      ' Tabｷｰはコントロール間移動なので処理をしない
      Return MyBase.ProcessCmdKey(msg, keyData)
    End If

    If TypeOf (Me.Columns(Me.CurrentCell.ColumnIndex)) Is DataGridViewButtonColumn Then
      ' セルがボタンの場合
      ' 本来はEnter、↑、↓キーでClickイベントが発生する。←キーではイベントは発生しない。
      ' EnterキーでTrueを返すようにしていると、Enterキーを押してもClickイベントが発生しない
      ' そこで以下の処理を追加する
      If (keyData And Keys.KeyCode) = Keys.Enter Then
        ' Enterキーで
        Return MyBase.ProcessCmdKey(msg, keyData)
      ElseIf (keyData And Keys.KeyCode) = Keys.Up Then
        ' ↑キーは無効。移動できない　
        Return True
      ElseIf (keyData And Keys.KeyCode) = Keys.Down Then
        ' ↓キーは無効。移動できない
        Return True
      End If
    End If

    Dim intCol As Integer = Me.CurrentCell.ColumnIndex
    Dim intRow As Integer = Me.CurrentCell.RowIndex

    If (keyData And Keys.KeyCode) = Keys.Enter Then NextStep = Tsugi.MAE
    If (keyData And Keys.KeyCode) = Keys.Up Then NextStep = Tsugi.UE
    If (keyData And Keys.KeyCode) = Keys.Down Then NextStep = Tsugi.SHITA

    If Me.IsCurrentCellInEditMode = True Then
      ' 編集中の場合
      If TypeOf Me.EditingControl Is DataGridViewTextBoxEditingControl Then
        ' テキストボックスで編集中の↑↓キーは処理をする（移動する）

      ElseIf TypeOf Me.EditingControl Is DataGridViewComboBoxEditingControl Then
        ' コンボボックスで編集中の↑↓キーは処理をしない
        If NextStep = Tsugi.UE Then NextStep = Tsugi.NASHI
        If NextStep = Tsugi.SHITA Then NextStep = Tsugi.NASHI
      End If
    Else
      '　編集中でない場合あるいは編集中でも、チェックボックス、ボタンコントロールの場合
      If (keyData And Keys.KeyCode) = Keys.Right Then NextStep = Tsugi.MIGI
      If (keyData And Keys.KeyCode) = Keys.Left Then NextStep = Tsugi.HIDARI
    End If

    If NextStep = Tsugi.NASHI Then
      Return MyBase.ProcessCmdKey(msg, keyData)
    End If

    ' **************************************************************************************************************
    ' セルが移動する場合は、イベントを発生させる
    Dim e As New System.Windows.Forms.DataGridViewCellEventArgs(Me.CurrentCell.ColumnIndex, Me.CurrentCell.RowIndex)
    RaiseEvent CellMove(Me, e)
    ' **************************************************************************************************************

    Dim OldCell As DataGridViewCell = Me.CurrentCell
    Dim NextCell As DataGridViewCell = Nothing

    Select Case NextStep
      Case Tsugi.MAE, Tsugi.MIGI
        '→
        Do
          If intCol = Me.ColumnCount - 1 Then
            If NextStep = Tsugi.MIGI Then Exit Do
            If intRow = Me.RowCount - 1 Then Exit Do
            intRow += 1
            intCol = -1
          End If
          intCol += 1
          If Me.Item(intCol, intRow).Visible = True Then
            If Me.Item(intCol, intRow).ReadOnly = False Then
              If Not (TypeOf Me.Columns(intCol) Is DataGridViewButtonColumn) Or (IsButtonVialbe = True) Then
                NextCell = Me.Item(intCol, intRow)
                Exit Do
              End If
            End If
          End If
        Loop

      Case Tsugi.HIDARI
        ' ←
        Do
          If intCol = 0 Then
            If NextStep = Tsugi.HIDARI Then Exit Do
            If intRow = 0 Then Exit Do
            intRow -= 1
            intCol = Me.ColumnCount
          End If
          intCol -= 1
          If Me.Item(intCol, intRow).Visible = True Then
            If Me.Item(intCol, intRow).ReadOnly = False Then
              If Not (TypeOf Me.Columns(intCol) Is DataGridViewButtonColumn) Or (IsButtonVialbe = True) Then
                NextCell = Me.Item(intCol, intRow)
                Exit Do
              End If
            End If
          End If
        Loop

      Case Tsugi.UE
        ' ↑
        Return MyBase.ProcessCmdKey(msg, keyData)

      Case Tsugi.SHITA
        ' ↓
        Return MyBase.ProcessCmdKey(msg, keyData)

    End Select

    If NextCell Is Nothing Then
      ' 移動しない場合：グリッドの先頭または最終セルの場合 
      ' -> フォーム側がグリッドのKeyDownを検出できるようにする
      If Me.IsCurrentCellInEditMode = True AndAlso TypeOf Me.EditingControl Is DataGridViewTextBoxEditingControl Then
        ' セルに入力してEnterキーなどを押した場合
        Me.EndEdit()
        If NextStep = Tsugi.MAE Or NextStep = Tsugi.MIGI Then
          OnKeyDown(New System.Windows.Forms.KeyEventArgs(Keys.Tab))
        Else
          OnKeyDown(New System.Windows.Forms.KeyEventArgs(Keys.Tab Or Keys.Shift))
        End If
        ' Trueか基本クラスのプロセスを実行する。
        ' Trueは処理なし。基本クラスのプロセスを実行するのも処理をしないので同じ結果になる。
        Return MyBase.ProcessCmdKey(msg, keyData)
      Else
        If NextStep = Tsugi.MAE Then
          ' Enterキーの場合は抜ける <- フォーム側のKeyDownイベントで処理
          Return MyBase.ProcessCmdKey(msg, keyData)
        Else
          ' Right、Leftキーの場合は処理なし。
          Return True
        End If
      End If
    End If

    ' 移動する場合する場合
    ' 移動がキャンセルされた場合（RowValidatingでe.Canceld = True）は例外が発生する。
    Try
      Me.CurrentCell = NextCell
      Return True
    Catch ex As Exception
      Return True
    Finally
    End Try
  End Function

  Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
    ' 編集モードで無い状態で、Deleteキーを押した場合の処理
    If Me Is Nothing OrElse Me.ReadOnly = True OrElse Me.CurrentCell Is Nothing Then
      MyBase.OnKeyDown(e)
      Return
    End If

    If e.KeyCode = Keys.Insert AndAlso Me.CurrentRow.Selected = True Then
      ' 行の挿入
      ' 行を追加してからValueを移動すると最下行に新規行を確保したままカレント行の前に１行挿入できる
      Me.Rows.Add()

      Dim i As Integer
      Dim j As Integer
      For i = Me.Rows.Count - 1 To Me.CurrentRow.Index + 1 Step -1
        For j = 0 To Me.ColumnCount - 1
          Me.Rows(i).Cells(j).Value = Me.Rows(i - 1).Cells(j).Value
        Next
      Next
      i = Me.CurrentRow.Index
      For j = 0 To Me.ColumnCount - 1
        Me.Rows(i).Cells(j).Value = Nothing
      Next

      MyBase.OnKeyDown(e)
      Return
    End If

    If e.KeyCode <> Keys.Delete Then
      MyBase.OnKeyDown(e)
      Return
    End If

    If TypeOf Me.Columns(Me.CurrentCell.ColumnIndex) Is DataGridViewTextBoxColumn Then
      ' テキストボックスの場合
      If Me.CurrentRow.Selected = False Then
        If Me.IsCurrentCellInEditMode = False Then
          ' 行全体が選択されていない＋編集モードでない

          ' クリアーする前に現在のセルの値をクリップボードにコピーする
          Clipboard.SetDataObject(Me.GetClipboardContent())
          ' セルの内容をクリアーして編集モードにする
          Me.CurrentCell.Value = Nothing
          Me.BeginEdit(True)

        ElseIf Me.IsCurrentCellInEditMode = True Then
          ' 編集モードではこのメソッドは発生しない　-> 削除してもEscapeKeyで元に戻せる

          ' 行全体が選択されていない＋編集モード
          e.Handled = True
        End If
      End If
    End If

    If TypeOf Me.Columns(Me.CurrentCell.ColumnIndex) Is DataGridViewComboBoxColumn Then
      ' コンボボックスの場合
      ' 行全体が選択されてDeleteｷｰが押された場合は、処理をしない -> 行削除を行う
      If Me.CurrentRow.Selected = False AndAlso Me.IsCurrentCellInEditMode = False Then
        Me.BeginEdit(True)
        If ComboEditCtrl IsNot Nothing Then
          ComboEditCtrl.SelectedItem = ""
        End If
      End If
    End If

    'If Me.CurrentRow IsNot Nothing AndAlso Me.ReadOnly = False Then
    '  ' 行全体が選択されてDeleteｷｰが押された場合
    '  If Me.CurrentRow.Selected = True AndAlso Me.IsCurrentCellInEditMode = False Then
    '    Dim e2 As DataGridViewCellCancelEventArgs = New DataGridViewCellCancelEventArgs(0, Me.CurrentRow.Index)
    '    RaiseEvent RowRemoving(Me, e2)
    '    If e2.Cancel = True Then
    '      e.Handled = True
    '      Return
    '    End If
    '  End If
    'End If


    MyBase.OnKeyDown(e)
  End Sub

  ' *************************************************
  ' グリッドのテキストの編集コントロールで
  ' 数値以外の文字入力をカットする
  ' *************************************************
  Private TextEditCtrl As DataGridViewTextBoxEditingControl
  Private MaskedEditCtrl As DataGridViewMaskedTextBoxEditingControl
  Private ComboEditCtrl As DataGridViewComboBoxEditingControl

  ' 編集モードに入るとき
  Protected Overrides Sub OnEditingControlShowing(ByVal e As DataGridViewEditingControlShowingEventArgs)

    If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
      'mEditingState = 1
      TextEditCtrl = CType(e.Control, DataGridViewTextBoxEditingControl)
      AddHandler TextEditCtrl.KeyPress, AddressOf Cell_KeyPress
    End If
    If TypeOf e.Control Is DataGridViewComboBoxEditingControl Then
      'mEditingState = 2
      ComboEditCtrl = CType(e.Control, DataGridViewComboBoxEditingControl)
      AddHandler ComboEditCtrl.SelectedIndexChanged, AddressOf Cell_SelectedIndexChanged
    End If
    If TypeOf e.Control Is DataGridViewMaskedTextBoxEditingControl Then
      MaskedEditCtrl = CType(e.Control, DataGridViewMaskedTextBoxEditingControl)
    End If

    mChanged = True             ' <---------- 変更済みスウィッチを立てる
    MyBase.OnEditingControlShowing(e)
  End Sub

  ' 編集モードを終了するとき
  Protected Overrides Sub OnCellEndEdit(ByVal e As DataGridViewCellEventArgs)
    If TextEditCtrl IsNot Nothing Then
      ' CheckBoxのようにOnEditingControlShowingが発生しないのに、OnCellEndEditが発生する場合がある
      RemoveHandler TextEditCtrl.KeyPress, AddressOf Cell_KeyPress
    End If

    If ComboEditCtrl IsNot Nothing Then
      RemoveHandler ComboEditCtrl.SelectedIndexChanged, AddressOf Cell_SelectedIndexChanged
    End If

    'mEditingState = 0

    Dim oCell As DataGridViewCell = Me.CurrentCell

    If oCell.Value IsNot Nothing AndAlso oCell.Value.ToString().Length > 0 Then

      Dim strFormat As String = Me.CurrentCell.Style.Format
      If strFormat = "" Then
        strFormat = Me.Columns(Me.CurrentCell.ColumnIndex).DefaultCellStyle.Format
      End If

      Select Case strFormat
        Case "N0"
          oCell.Value = Format(CDbl(oCell.Value), "###,###,##0")
        Case "N1"
          oCell.Value = Format(CDbl(oCell.Value), "###,###,##0.0")
        Case "N2"
          oCell.Value = Format(CDbl(oCell.Value), "###,###,##0.00")
        Case "N3"
          oCell.Value = Format(CDbl(oCell.Value), "###,###,##0.000")
        Case "N4"
          oCell.Value = Format(CDbl(oCell.Value), "###,###,##0.0000")
          'Case "d"
          '  Dim strText As String = oCell.Value.ToString()
          '  If strText.IndexOf("/") = -1 Then
          '    If Strings.Len(strText) < 3 Then
          '      strText = Strings.Left(Format(Now(), "yyyyMMdd"), 6) & Strings.Right("00" & strText, 2)
          '    ElseIf Strings.Len(strText) < 5 Then
          '      strText = Strings.Left(Format(Now(), "yyyyMMdd"), 4) & Strings.Right("0000" & strText, 4)
          '    End If
          '    If IsNumeric(strText) Then
          '      strText = Format(CLng(strText), "####/##/##")
          '    End If
          '  End If
          '  If IsDate(strText) Then
          '    Dim oDate As Date = CDate(strText)
          '    oCell.Value = oDate.Year * 10000 + oDate.Month * 100 + oDate.Day
          '  Else
          '    oCell.Value = ""
          '  End If

      End Select
    End If
    MyBase.OnCellEndEdit(e)
  End Sub

  ' 編集モードのキープレスイベント処理 -> 数字以外はカット
  Private Sub Cell_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
    If Me.CurrentCell Is Nothing Then Return

    If e.KeyChar = Chr(Keys.Back) Then
      MyBase.OnKeyPress(e)
      Return
    End If

    Select Case Me.Columns(Me.CurrentCell.ColumnIndex).DefaultCellStyle.Format.ToString()
      Case "N0"
        ' 少数点なし
        If e.KeyChar <> "-"c And (e.KeyChar < "0"c Or e.KeyChar > "9"c) Then
          e.Handled = True
        End If

      Case "N1", "N2", "N3", "N4"
        ' 小数点あり
        If e.KeyChar <> "-"c And e.KeyChar <> "."c And (e.KeyChar < "0"c Or e.KeyChar > "9"c) Then
          e.Handled = True
        End If
    End Select
    MyBase.OnKeyPress(e)
  End Sub

  ' グリッド内のコントロールのイベントを拾って、アプリ側にイベントを発生させる
  ' ComboEditCtrl.SelectedIndexChangedイベントのイベントハンドラ内でRaiseEventを実行する
  '
  Private Sub Cell_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    ' typeof sender is DataGridViewComboBoxEditingControl

    If sender Is Nothing Then Return
    If Me.CurrentCell Is Nothing Then Return
    If Me.ComboEditCtrl Is Nothing Then Return

    ' senderに別のｵﾌﾞｼﾞｪｸﾄを指定できる。eに別のｵﾌﾞｼﾞｪｸﾄﾄを指定できる
    Dim e2 As New ExDataGridViewCellEventArgs(Me.CurrentCell.ColumnIndex, Me.CurrentCell.RowIndex, Me.ComboEditCtrl.Text)
    RaiseEvent SelectedIndexChanged(Me, e2)
  End Sub

  'Protected Friend Overrides Sub OnCellPainting(ByVal e As DataGridViewCellPaintingEventArgs)
  Protected Overrides Sub OnCellPainting(ByVal e As DataGridViewCellPaintingEventArgs)
    MyBase.OnCellPainting(e)
    Return

    'Dim charBrush As SolidBrush
    'Dim foreBrush As SolidBrush
    'Dim backBrush As SolidBrush
    'Dim redBrush As New SolidBrush(Color.Red)


    'If (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
    '  foreBrush = New SolidBrush(e.CellStyle.SelectionForeColor)
    '  backBrush = New SolidBrush(e.CellStyle.SelectionBackColor)
    'Else
    '  foreBrush = New SolidBrush(e.CellStyle.ForeColor)
    '  backBrush = New SolidBrush(e.CellStyle.BackColor)
    'End If

    ' ①
    ' 背景：指定された四角を塗りつぶす
    ' e.CellBoundsはRectangleを返す
    'e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

    ' ②
    ' ClipBounds プロパティは、再描画が必要な DataGridView の領域を表します。
    ' 描画する部分を指定する DataGridViewPaintParts 値のビットごとの組み合わせ。
    ' DataGridViewPaintPartsは列挙体
    ' Background セルの背景を描画する必要があります。 
    ' Focus セルの周囲にフォーカスを示す四角形を描画する必要があります。 
    'e.Paint(e.ClipBounds, DataGridViewPaintParts.Border Or DataGridViewPaintParts.Focus)

    'If e.Value IsNot Nothing AndAlso CStr(e.Value).Length > 0 Then
    '  Dim diffX As Integer = 0
    '  For Each c As Char In CStr(e.Value).ToCharArray()
    '    ' "う"の場合のみ赤文字にする
    '    If c = "る"c Then
    '      charBrush = redBrush
    '    Else
    '      charBrush = foreBrush
    '    End If

    '    ' １文字描画
    '    e.Graphics.DrawString(CStr(c), _
    '                          e.CellStyle.Font, _
    '                          charBrush, _
    '                          e.CellBounds.X + diffX, _
    '                          e.CellBounds.Y + 2, _
    '                          StringFormat.GenericDefault)

    '    ' 描画した文字幅を取得して加算
    '    diffX += GetCharWidth(e.Graphics, _
    '                          e.CellStyle.Font, _
    '                          New RectangleF(e.CellBounds.X, e.CellBounds.Y, e.CellBounds.Width, e.CellBounds.Height), _
    '                          CStr(c))
    '  Next

    '  foreBrush.Dispose()
    '  backBrush.Dispose()
    '  redBrush.Dispose()

    '  e.Handled = True
    'End If

    'MyBase.OnCellPainting(e)
  End Sub

  Private Function GetCharWidth(ByVal g As Graphics, ByVal f As Font, ByVal layout As RectangleF, ByVal str As String) As Integer
    ' CharacterRange構造体。初期化する場合は{}で括る。
    'Dim charRanges() As CharacterRange = {New CharacterRange(0, str.Length), New CharacterRange(0, 1)}
    Dim charRanges() As CharacterRange = {New CharacterRange(0, str.Length)}
    Dim strFmt As New StringFormat

    ' ②　StringFormatの初期化
    strFmt.SetMeasurableCharacterRanges(charRanges)

    ' ③　Regionの初期化
    Dim strRegions() As Region = g.MeasureCharacterRanges(str, f, layout, strFmt)

    ' ④　
    Dim rect As RectangleF = strRegions(0).GetBounds(g)

    Return rect.Width - 1

  End Function


  Protected Overrides Sub OnRowEnter(ByVal e As DataGridViewCellEventArgs)
    If e.RowIndex >= 0 Then
      For i As Integer = 0 To Me.Columns.Count - 1
        mItem(i) = Me.Rows(e.RowIndex).Cells(i).Value
      Next
    End If
    MyBase.OnRowEnter(e)
  End Sub

  Public Sub Recovery(ByRef rowIndex As Integer)
    For i As Integer = 0 To Me.Columns.Count - 1
      'Select Case Me.Columns(i).DefaultCellStyle.Format.ToString()
      '  Case "N0"
      '    ' 少数点なし
      '    If IsNumeric(mItem(i)) Then
      '      Me.Rows(rowIndex).Cells(i).Value = Format(CDbl(mItem(i)), "###,###,###")
      '    End If

      '  Case "N1"
      '    If IsNumeric(mItem(i)) Then
      '      Me.Rows(rowIndex).Cells(i).Value = Format(CDbl(mItem(i)), "###,###,##0.0")
      '    End If

      '  Case "N2"
      '    If IsNumeric(mItem(i)) Then
      '      Me.Rows(rowIndex).Cells(i).Value = Format(CDbl(mItem(i)), "###,###,##0.00")
      '    End If

      '  Case "N3"
      '    If IsNumeric(mItem(i)) Then
      '      Me.Rows(rowIndex).Cells(i).Value = Format(CDbl(mItem(i)), "###,###,##0.000")
      '    End If

      '  Case "N4"
      '    If IsNumeric(mItem(i)) Then
      '      Me.Rows(rowIndex).Cells(i).Value = Format(CDbl(mItem(i)), "###,###,##0.0000")
      '    End If

      'End Select

      Me.Rows(rowIndex).Cells(i).Value = mItem(i)
    Next
  End Sub

  'Public Property Changed() As Boolean
  '  Get
  '    Return mChanged
  '  End Get
  '  Set(ByVal value As Boolean)
  '    mChanged = value
  '  End Set
  'End Property

  Public Property IsButtonVialbe() As Boolean
    Get
      Return mIsButtonViable
    End Get
    Set(ByVal value As Boolean)
      mIsButtonViable = value
    End Set
  End Property

  Protected Overrides Sub OnRowValidating(ByVal e As DataGridViewCellCancelEventArgs)
    If Me.ReadOnly Then
      MyBase.OnRowValidating(e)
      Return
    End If

    If mChanged = False Then
      ' 変更済みスウィッチがFalseの場合は、フォーム側でイベントが発生しないようにする。
      Return
    Else
      ' 変更済みスウィッチがTrueの場合は、Falseに戻し、イベントを発生させる
      mChanged = False
      MyBase.OnRowValidating(e)
    End If
  End Sub

  ' OnEneter -> OnGotFocus 
  ' OnLostFocus -> OnLeave
  Protected Overrides Sub OnEnter(ByVal e As EventArgs)
    ' 「変更されていない」に設定する
    mChanged = False
    MyBase.OnEnter(e)
  End Sub

  Protected Overrides Sub OnLeave(ByVal e As EventArgs)
    MyBase.OnLeave(e)
  End Sub

  Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
    ' 編集モードが終了したときにも発生する
    MyBase.OnGotFocus(e)
  End Sub

  Protected Overrides Sub OnLostFocus(ByVal e As EventArgs)
    ' 編集モードに入るときにも発生する 
    MyBase.OnLostFocus(e)
  End Sub

  Protected Overrides Sub OnCellEnter(ByVal e As DataGridViewCellEventArgs)
    'mEditingState = 0
    mCurrentValue = ""
    If Me.Rows(e.RowIndex).Cells(e.ColumnIndex).Value IsNot Nothing Then
      mCurrentValue = Me.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()
    End If
    MyBase.OnCellEnter(e)
  End Sub

  Protected Overrides Sub OnCellLeave(ByVal e As DataGridViewCellEventArgs)
    MyBase.OnCellLeave(e)
  End Sub

  Public Property CurrentValue() As String
    Get
      If TextEditCtrl IsNot Nothing Then
        Return TextEditCtrl.Text
      End If
      If ComboEditCtrl IsNot Nothing Then
        Return ComboEditCtrl.Text
      End If
      If MaskedEditCtrl IsNot Nothing Then
        Return MaskedEditCtrl.Text
      End If

      Return mCurrentValue
    End Get

    Set(ByVal value As String)
      If TextEditCtrl IsNot Nothing Then
        TextEditCtrl.Text = value
        TextEditCtrl.SelectAll()
      End If

      If ComboEditCtrl IsNot Nothing Then
        If ComboEditCtrl.Visible = False Then ComboEditCtrl.Visible = True
        ComboEditCtrl.Text = value
        ComboEditCtrl.SelectAll()
      End If

      If MaskedEditCtrl IsNot Nothing Then
        MaskedEditCtrl.Text = value
        MaskedEditCtrl.SelectAll()
      End If

      mCurrentValue = value
    End Set
  End Property

  Public ReadOnly Property CurrentValue(ByVal RowIndex As Integer, ByVal ColumnIndex As Integer) As String
    Get
      If RowIndex > Me.Rows.Count - 1 Then Return ""
      If ColumnIndex > Me.Columns.Count - 1 Then Return ""

      If Me.CurrentCell IsNot Nothing AndAlso Me.CurrentCell.RowIndex = RowIndex AndAlso Me.CurrentCell.ColumnIndex = ColumnIndex Then
        ' 編集中のセルの値を参照する場合
        If TextEditCtrl IsNot Nothing Then
          Return TextEditCtrl.Text
        End If
        If ComboEditCtrl IsNot Nothing Then
          Return ComboEditCtrl.Text
        End If
        If MaskedEditCtrl IsNot Nothing Then
          Return MaskedEditCtrl.Text
        End If
        Return mCurrentValue
      End If

      ' 編集中以外のセルの値
      Return gStringRet(Me.Rows(RowIndex).Cells(ColumnIndex).Value)
    End Get
  End Property


  'Dim mCol As Integer = -1
  'Dim mRow As Integer = -1
  'Dim mState As Integer = 0
  'Dim mNewCol As Integer
  'Dim mNewRow As Integer

  ' 点線で囲まれた部分
  'Protected Overrides Function SetCurrentCellAddressCore(ByVal columnIndex As Integer, ByVal rowIndex As Integer, ByVal setAnchorCellAddress As Boolean, ByVal validateCurrentCell As Boolean, ByVal throughMouseClick As Boolean) As Boolean
  '  If columnIndex = -1 Then
  '    mCol = -1
  '    mRow = -1
  '    mState = 0
  '    Return MyBase.SetCurrentCellAddressCore(columnIndex, rowIndex, setAnchorCellAddress, validateCurrentCell, throughMouseClick)
  '  End If

  '  Dim blnDirect As Boolean = False

  '  If (mState And 2) = 0 Then
  '    mNewCol = columnIndex
  '    mNewRow = rowIndex

  '    blnDirect = ((mNewRow * 100 + mNewCol) >= (mRow * 100 + mCol))
  '    NextCell(mNewCol, mNewRow, blnDirect)
  '    If mNewCol = -1 And mNewRow = -1 Then
  '      mNewCol = mCol
  '      mNewRow = mRow
  '    End If
  '    mCol = mNewCol
  '    mRow = mNewRow
  '    mState = (mState Or 1)
  '  Else
  '    mState = 0
  '  End If

  '  Return MyBase.SetCurrentCellAddressCore(mNewCol, mNewRow, setAnchorCellAddress, validateCurrentCell, throughMouseClick)
  'End Function

  ' 反転表示
  'Protected Overrides Sub SetSelectedCellCore(ByVal columnIndex As Integer, ByVal rowIndex As Integer, ByVal selected As Boolean)
  '  Dim blnDirect As Boolean

  '  If (mState And 1) = 0 Then
  '    mNewCol = columnIndex
  '    mNewRow = rowIndex
  '    blnDirect = ((mNewRow * 100 + mNewCol) >= (mRow * 100 + mCol))

  '    NextCell(mNewCol, mNewRow, blnDirect)
  '    If mNewCol = -1 And mNewRow = -1 Then
  '      mNewCol = mCol
  '      mNewRow = mRow
  '    End If

  '    mCol = mNewCol
  '    mRow = mNewRow

  '    mState = (mState Or 2)
  '  Else
  '    mState = 0
  '  End If

  '  MyBase.SetSelectedCellCore(mNewCol, mNewRow, selected)
  'End Sub

  'Private Sub NextCell(ByRef NewCol As Integer, ByRef NewRow As Integer, ByVal blnDirect As Boolean)
  '  Dim intCol As Integer = NewCol
  '  Dim intRow As Integer = NewRow

  '  Do
  '    If Me.ReadOnly = True Then Exit Do
  '    If intCol < 0 Or intRow < 0 Then Exit Do

  '    If Me.Item(intCol, intRow).Visible = True AndAlso Me.Item(intCol, intRow).ReadOnly = False Then
  '      NewCol = intCol
  '      NewRow = intRow
  '      Exit Do
  '    End If

  '    If blnDirect = True Then
  '      intCol += 1
  '      If intCol > Me.ColumnCount - 1 Then
  '        intCol = 0
  '        intRow += 1
  '      End If
  '    Else
  '      intCol -= 1
  '      If intCol < 0 Then
  '        intCol = Me.ColumnCount - 1
  '        intRow -= 1
  '      End If
  '    End If
  '    If intRow > Me.RowCount - 1 Or intRow < 0 Then
  '      NewCol = -1
  '      NewRow = -1
  '      Exit Do
  '    End If
  '  Loop

  'End Sub


  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
End Class

Public Class ExDataGridViewCellEventArgs
  Inherits System.Windows.Forms.DataGridViewCellEventArgs

  Private mValue As Object

  Public ReadOnly Property Value() As Object
    Get
      Return mValue
    End Get
  End Property

  Public Sub New(ByVal columnIndex As Integer, ByVal rowIndex As Integer, ByVal value As Object)
    MyBase.New(columnIndex, rowIndex)
    mValue = value
  End Sub
End Class



Public Class DataGridViewMaskedTextBoxColumn
  Inherits DataGridViewColumn

  'CellTemplateとするDataGridViewMaskedTextBoxCellオブジェクトを指定して
  '基本クラスのコンストラクタを呼び出す
  Public Sub New()
    MyBase.New(New DataGridViewMaskedTextBoxCell())
  End Sub

  Private maskValue As String = ""

  Public Property Mask() As String
    Get
      Return Me.maskValue
    End Get
    Set(ByVal value As String)
      Me.maskValue = value
    End Set
  End Property

  '新しいプロパティを追加しているため、
  ' Cloneメソッドをオーバーライドする必要がある
  Public Overrides Function Clone() As Object
    Dim col As DataGridViewMaskedTextBoxColumn = CType(MyBase.Clone(), DataGridViewMaskedTextBoxColumn)
    col.Mask = Me.Mask
    Return col
  End Function

  'CellTemplateの取得と設定
  Public Overrides Property CellTemplate() As DataGridViewCell
    Get
      Return MyBase.CellTemplate
    End Get
    Set(ByVal value As DataGridViewCell)
      'DataGridViewMaskedTextBoxCellしか
      ' CellTemplateに設定できないようにする
      If Not TypeOf value Is DataGridViewMaskedTextBoxCell Then
        Throw New InvalidCastException( _
            "DataGridViewMaskedTextBoxCellオブジェクトを" + _
            "指定してください。")
      End If
      MyBase.CellTemplate = value
    End Set
  End Property
End Class

Public Class DataGridViewMaskedTextBoxCell
  Inherits DataGridViewTextBoxCell

  Private mMask As String = ""
  Private mAsciiOnly As Boolean = True
  Private mImeMode As ImeMode = ImeMode.NoControl

  'コンストラクタ
  Public Sub New()
  End Sub

  Public Property Mask() As String
    Get
      Return Me.mMask
    End Get
    Set(ByVal value As String)
      Me.mMask = value
    End Set
  End Property

  Public Property AsciiOnly() As Boolean
    Get
      Return mAsciiOnly
    End Get
    Set(ByVal value As Boolean)
      mAsciiOnly = value
    End Set
  End Property

  Public Property ImeMode() As ImeMode
    Get
      Return mImeMode
    End Get
    Set(ByVal value As ImeMode)
      mImeMode = value
    End Set
  End Property

  '編集コントロールを初期化する
  '編集コントロールは別のセルや列でも使いまわされるため、初期化の必要がある
  Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
                                                ByVal initialFormattedValue As Object, _
                                                ByVal dataGridViewCellStyle As DataGridViewCellStyle)

    MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

    '編集コントロールの取得
    'Dim maskedBox As DataGridViewMaskedTextBoxEditingControl = Me.DataGridView.EditingControl

    'If Not (maskedBox Is Nothing) Then
    '  'Textを設定
    '  maskedBox.Text = If(Me.Value Is Nothing, "", Me.Value.ToString())

    '  'カスタム列のプロパティを反映させる
    '  Dim column As DataGridViewMaskedTextBoxColumn = Me.OwningColumn
    '  If Not (column Is Nothing) Then
    '    maskedBox.Mask = column.Mask
    '  End If
    'End If

    Dim editCtl As DataGridViewMaskedTextBoxEditingControl = CType(DataGridView.EditingControl, DataGridViewMaskedTextBoxEditingControl)
    editCtl.Text = Me.Value
    editCtl.Mask = Me.mMask
    editCtl.AsciiOnly = Me.AsciiOnly
    editCtl.ImeMode = Me.ImeMode

    ' マスクの設定
    If TypeOf (Me.OwningColumn) Is DataGridViewMaskedTextBoxColumn Then
      ' コラムがマスクテキストボックスの場合
      editCtl.Mask = DirectCast(Me.OwningColumn, DataGridViewMaskedTextBoxColumn).Mask
    Else
      ' コラムのテキストボックスの場合
      editCtl.Mask = Me.mMask
    End If

    ' MaxInputLengthの設定
    'DirectCast(Me.OwningColumn, DataGridViewColumn).

  End Sub

  '編集コントロールの型を指定する
  Public Overrides ReadOnly Property EditType() As Type
    Get
      Return GetType(DataGridViewMaskedTextBoxEditingControl)
    End Get
  End Property

  'セルの値のデータ型を指定する
  'ここでは、Object型とする
  '基本クラスと同じなので、オーバーライドの必要なし
  Public Overrides ReadOnly Property ValueType() As Type
    Get
      Return GetType(Object)
    End Get
  End Property

  '新しいレコード行のセルの既定値を指定する
  Public Overrides ReadOnly Property DefaultNewRowValue() As Object
    Get
      Return MyBase.DefaultNewRowValue
    End Get
  End Property
End Class

Public Class DataGridViewMaskedTextBoxEditingControl
  Inherits MaskedTextBox
  Implements IDataGridViewEditingControl

  '編集コントロールが表示されているDataGridView
  Private dataGridView As DataGridView
  '編集コントロールが表示されている行
  Private rowIndex As Integer
  '編集コントロールの値とセルの値が違うかどうか
  Private valueChanged As Boolean

  'コンストラクタ
  Public Sub New()
    Me.TabStop = False
  End Sub

  '編集コントロールで変更されたセルの値
  Public Function GetEditingControlFormattedValue(ByVal context As DataGridViewDataErrorContexts) As Object _
      Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

    Return Me.Text
  End Function

  '編集コントロールで変更されたセルの値
  Public Property EditingControlFormattedValue() As Object _
      Implements IDataGridViewEditingControl.EditingControlFormattedValue
    Get
      Return Me.GetEditingControlFormattedValue( _
          DataGridViewDataErrorContexts.Formatting)
    End Get
    Set(ByVal value As Object)
      Me.Text = CStr(value)
    End Set
  End Property

  'セルスタイルを編集コントロールに適用する
  '編集コントロールの前景色、背景色、フォントなどをセルスタイルに合わせる
  Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As DataGridViewCellStyle) _
      Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

    Me.Font = dataGridViewCellStyle.Font
    Me.ForeColor = dataGridViewCellStyle.ForeColor
    Me.BackColor = dataGridViewCellStyle.BackColor
    Select Case dataGridViewCellStyle.Alignment
      Case DataGridViewContentAlignment.BottomCenter, _
              DataGridViewContentAlignment.MiddleCenter, _
              DataGridViewContentAlignment.TopCenter
        Me.TextAlign = HorizontalAlignment.Center
      Case DataGridViewContentAlignment.BottomRight, _
              DataGridViewContentAlignment.MiddleRight, _
              DataGridViewContentAlignment.TopRight
        Me.TextAlign = HorizontalAlignment.Right
      Case Else
        Me.TextAlign = HorizontalAlignment.Left
    End Select
  End Sub

  '編集するセルがあるDataGridView
  Public Property EditingControlDataGridView() As DataGridView _
      Implements IDataGridViewEditingControl.EditingControlDataGridView
    Get
      Return Me.dataGridView
    End Get
    Set(ByVal value As DataGridView)
      Me.dataGridView = value
    End Set
  End Property

  '編集している行のインデックス
  Public Property EditingControlRowIndex() As Integer _
      Implements IDataGridViewEditingControl.EditingControlRowIndex
    Get
      Return Me.rowIndex
    End Get
    Set(ByVal value As Integer)
      Me.rowIndex = value
    End Set
  End Property

  '値が変更されたかどうか
  '編集コントロールの値とセルの値が違うかどうか
  Public Property EditingControlValueChanged() As Boolean _
      Implements IDataGridViewEditingControl.EditingControlValueChanged
    Get
      Return Me.valueChanged
    End Get
    Set(ByVal value As Boolean)
      Me.valueChanged = value
    End Set
  End Property

  '指定されたキーをDataGridViewが処理するか、編集コントロールが処理するか
  'Trueを返すと、編集コントロールが処理する
  'dataGridViewWantsInputKeyがTrueの時は、DataGridViewが処理できる
  Public Function EditingControlWantsInputKey(ByVal keyData As Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
      Implements IDataGridViewEditingControl.EditingControlWantsInputKey

    'Keys.Left、Right、Home、Endの時は、Trueを返す
    'このようにしないと、これらのキーで別のセルにフォーカスが移ってしまう
    Select Case keyData And Keys.KeyCode
      Case Keys.Right, Keys.End, Keys.Left, Keys.Home
        Return True
      Case Else
        Return False
    End Select
  End Function

  'マウスカーソルがEditingPanel上にあるときのカーソルを指定する
  'EditingPanelは編集コントロールをホストするパネルで、
  '編集コントロールがセルより小さいとコントロール以外の部分がパネルとなる
  Public ReadOnly Property EditingPanelCursor() As Cursor _
      Implements IDataGridViewEditingControl.EditingPanelCursor
    Get
      Return MyBase.Cursor
    End Get
  End Property

  'コントロールで編集する準備をする
  'テキストを選択状態にしたり、挿入ポインタを末尾にしたりする
  Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
      Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

    If selectAll Then
      '選択状態にする
      Me.SelectAll()
    Else
      '挿入ポインタを末尾にする
      'Me.SelectionStart = Me.TextLength
      Me.SelectionStart = 0
    End If
  End Sub

  '値が変更した時に、セルの位置を変更するかどうか
  '値が変更された時に編集コントロールの大きさが変更される時はTrue
  Public ReadOnly Property RepositionEditingControlOnValueChange() As Boolean _
      Implements _
          IDataGridViewEditingControl.RepositionEditingControlOnValueChange
    Get
      Return False
    End Get
  End Property

  '値が変更された時
  Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
    MyBase.OnTextChanged(e)
    '値が変更されたことをDataGridViewに通知する
    Me.valueChanged = True
    Me.dataGridView.NotifyCurrentCellDirty(True)
  End Sub

  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
End Class


'Public Class CalendarColumn
'  Inherits DataGridViewColumn

'  Public Sub New()
'    MyBase.New(New CalendarCell())
'  End Sub

'  Public Overrides Property CellTemplate() As DataGridViewCell
'    Get
'      Return MyBase.CellTemplate
'    End Get
'    Set(ByVal value As DataGridViewCell)

'      ' Ensure that the cell used for the template is a CalendarCell.
'      If Not (value Is Nothing) AndAlso Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) Then
'        Throw New InvalidCastException("Must be a CalendarCell")
'      End If

'      MyBase.CellTemplate = value

'    End Set
'  End Property

'End Class

'Public Class CalendarCell
'  Inherits DataGridViewTextBoxCell

'  Public Sub New()
'    ' Use the short date format.
'    Me.Style.Format = "d"
'  End Sub

'  Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As DataGridViewCellStyle)

'    ' Set the value of the editing control to the current cell value.
'    MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

'    Dim ctl As CalendarEditingControl = CType(DataGridView.EditingControl, CalendarEditingControl)

'    'ctl.Value = CType(Me.Value, DateTime)

'    If Me.Value Is Nothing Then
'      ' DateTimePickerのValueにはNothingは設定できない
'      ctl.Value = Now()
'      Return
'    End If

'    ctl.Value = If(IsDate(Me.Value), Me.Value, Nothing)

'  End Sub

'  Public Overrides ReadOnly Property EditType() As Type
'    Get
'      ' Return the type of the editing contol that CalendarCell uses.
'      Return GetType(CalendarEditingControl)
'    End Get
'  End Property

'  Public Overrides ReadOnly Property ValueType() As Type
'    Get
'      ' Return the type of the value that CalendarCell contains.
'      Return GetType(DateTime)
'    End Get
'  End Property

'  Public Overrides ReadOnly Property DefaultNewRowValue() As Object
'    Get
'      ' Use the current date and time as the default value.
'      Return DateTime.Now
'    End Get
'  End Property

'End Class

'Class CalendarEditingControl
'  Inherits DateTimePicker
'  Implements IDataGridViewEditingControl

'  Private dataGridViewControl As DataGridView
'  Private valueIsChanged As Boolean = False
'  Private rowIndexNum As Integer

'  Public Sub New()
'    Me.Format = DateTimePickerFormat.Short
'  End Sub

'  Public Property EditingControlFormattedValue() As Object _
'      Implements IDataGridViewEditingControl.EditingControlFormattedValue

'    Get
'      Return Me.Value.ToShortDateString()
'    End Get

'    Set(ByVal value As Object)
'      If TypeOf value Is [String] Then
'        Me.Value = DateTime.Parse(CStr(value))
'      End If
'    End Set

'  End Property

'  Public Function GetEditingControlFormattedValue(ByVal context As DataGridViewDataErrorContexts) As Object _
'      Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

'    Return Me.Value.ToShortDateString()

'  End Function

'  Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As DataGridViewCellStyle) _
'      Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

'    Me.Font = dataGridViewCellStyle.Font
'    Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
'    Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

'  End Sub

'  Public Property EditingControlRowIndex() As Integer _
'      Implements IDataGridViewEditingControl.EditingControlRowIndex

'    Get
'      Return rowIndexNum
'    End Get
'    Set(ByVal value As Integer)
'      rowIndexNum = value
'    End Set

'  End Property

'  Public Function EditingControlWantsInputKey(ByVal key As Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
'      Implements IDataGridViewEditingControl.EditingControlWantsInputKey

'    ' Let the DateTimePicker handle the keys listed.
'    Select Case key And Keys.KeyCode
'      Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
'          Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

'        Return True

'      Case Else
'        Return False
'    End Select

'  End Function

'  Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
'      Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

'    ' No preparation needs to be done.

'  End Sub

'  Public ReadOnly Property RepositionEditingControlOnValueChange() As Boolean Implements _
'      IDataGridViewEditingControl.RepositionEditingControlOnValueChange

'    Get
'      Return False
'    End Get

'  End Property

'  Public Property EditingControlDataGridView() As DataGridView _
'      Implements IDataGridViewEditingControl.EditingControlDataGridView

'    Get
'      Return dataGridViewControl
'    End Get
'    Set(ByVal value As DataGridView)
'      dataGridViewControl = value
'    End Set

'  End Property

'  Public Property EditingControlValueChanged() As Boolean _
'      Implements IDataGridViewEditingControl.EditingControlValueChanged

'    Get
'      Return valueIsChanged
'    End Get
'    Set(ByVal value As Boolean)
'      valueIsChanged = value
'    End Set

'  End Property

'  Public ReadOnly Property EditingControlCursor() As Cursor _
'      Implements IDataGridViewEditingControl.EditingPanelCursor

'    Get
'      Return MyBase.Cursor
'    End Get

'  End Property

'  Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)

'    ' Notify the DataGridView that the contents of the cell have changed.
'    valueIsChanged = True
'    Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
'    MyBase.OnValueChanged(eventargs)

'  End Sub

'End Class


