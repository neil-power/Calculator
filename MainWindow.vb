Public Class MainWindow
    Dim calculation As New ArrayList

    Private Function ArrayToString(ByVal arrayList)
        Dim builder As New System.Text.StringBuilder()
        For Each num In arrayList
            builder.Append(num.ToString)
        Next
        Return builder.ToString()
    End Function

    Private Sub ButtonEquals_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ButtonEquals.Click
        Dim finalCalculation() As String = CreateCalculation()
        Calculate(finalCalculation)
    End Sub

    Private Sub Calculate(ByVal finalCalculation As String())
        Dim integerTest As Integer
        Dim doubleTest As Double
        Dim result As Double
        Dim mathsOperators As Integer = 1
        Dim mathsOperator As String = ""
        For index = 0 To finalCalculation.Length - 1
            Dim num As String = finalCalculation(index)
            If (Integer.TryParse(num, integerTest) Or num = "." Or Double.TryParse(num, doubleTest)) And (mathsOperators Mod 2 = 0) = False Then
                result &= num
            ElseIf (num.ToString = "+") Or (num.ToString = "-") Or (num.ToString = "/") Or (num.ToString = "*") Then
                mathsOperators += 1
                mathsOperator = num
            ElseIf (Integer.TryParse(num, integerTest) Or num = "." Or Double.TryParse(num, doubleTest)) And mathsOperators Mod 2 = 0 Then
                Select Case mathsOperator
                    Case "+"
                        result = result + Convert.ToDouble(num)
                    Case "-"
                        result = result - Convert.ToDouble(num)
                    Case "*"
                        result = result * Convert.ToDouble(num)
                    Case "/"
                        result = result / Convert.ToDouble(num)
                    Case Else
                        MsgBox("Error?")
                End Select
                mathsOperators += 1
            End If
        Next
        UpdateDisplay(result)
        ClearAll(False)
    End Sub

    Private Sub ClearAll(ByVal resetDisplay As Boolean)
        calculation.Clear()
        If resetDisplay = True Then
            UpdateDisplay(calculation)
            Dim labelFont As Font
            labelFont = ResultLabel.Font
            ResultLabel.Font = New Font(labelFont.Name, 30)
        End If

    End Sub

    Private Function CreateCalculation()
        Dim finalCalculation(0 To calculation.Count) As String
        Dim integerTest As Integer
        Dim mathsOperators As Integer = 0
        For index = 0 To calculation.Count - 1
            Dim num As String = calculation(index)
            If Integer.TryParse(num, integerTest) Or num = "." Then
                finalCalculation(mathsOperators) &= num
            ElseIf (num.ToString = "+") Or (num.ToString = "-") Or (num.ToString = "/") Or (num.ToString = "*") Then
                mathsOperators += 1
                finalCalculation(mathsOperators) = num
                mathsOperators += 1
            End If
        Next
        Return RemoveEmptyValues(finalCalculation)
    End Function

    Private Sub MainWindow_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles MyBase.KeyPress
        Dim keyPressed As String = e.KeyChar.ToString
        Select Case e.KeyChar.ToString
            Case 0
                calculation.Add(e.KeyChar.ToString)
            Case 1
                calculation.Add(e.KeyChar.ToString)
            Case 2
                calculation.Add(e.KeyChar.ToString)
            Case 3
                calculation.Add(e.KeyChar.ToString)
            Case 4
                calculation.Add(e.KeyChar.ToString)
            Case 5
                calculation.Add(e.KeyChar.ToString)
            Case 6
                calculation.Add(e.KeyChar.ToString)
            Case 7
                calculation.Add(e.KeyChar.ToString)
            Case 8
                calculation.Add(e.KeyChar.ToString)
            Case 9
                calculation.Add(e.KeyChar.ToString)
            Case "+"
                calculation.Add(e.KeyChar.ToString)
            Case "-"
                calculation.Add(e.KeyChar.ToString)
            Case "*"
                calculation.Add(e.KeyChar.ToString)
            Case "/"
                calculation.Add(e.KeyChar.ToString)
            Case "."
                calculation.Add(e.KeyChar.ToString)
                'Case "="
                '    ButtonEquals_Click(ButtonEquals, "")
            Case Asc(8)
                If calculation.Count > 0 Then
                    calculation.RemoveAt(calculation.Count - 1)
                End If
            Case "?"
                Dim randInt As New Random
                calculation.Add(randInt.Next(1, 999))
            Case Else
                Exit Select
        End Select
        UpdateDisplay(calculation)
    End Sub

    Private Sub MainWindow_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ResultLabel.RightToLeft = RightToLeft.No
        Dim labelSize As New Size With {.Height = 45, .Width = 255}
        ResultLabel.MinimumSize = labelSize
        ResultLabel.MaximumSize = labelSize
        Me.KeyPreview = True
        ClearAll(True)
    End Sub

#Disable Warning CC0057 ' Unused parameters.
    Private Sub NumericButton_Click(ByVal sender As Button, ByVal e As EventArgs) Handles ButtonZero.Click, ButtonOne.Click, ButtonTwo.Click,
        ButtonThree.Click, ButtonFour.Click, ButtonFive.Click, ButtonSix.Click, ButtonSeven.Click, ButtonEight.Click, ButtonNine.Click,
        ButtonAdd.Click, ButtonSubtract.Click, ButtonMultiply.Click, ButtonDivide.Click, ButtonDecimalPoint.Click, ButtonCE.Click, MysteryButton.Click, ButtonC.Click
#Enable Warning CC0057 ' Unused parameters.
        Select Case sender.Name
            Case ButtonZero.Name
                calculation.Add(0)
            Case ButtonOne.Name
                calculation.Add(1)
            Case ButtonTwo.Name
                calculation.Add(2)
            Case ButtonThree.Name
                calculation.Add(3)
            Case ButtonFour.Name
                calculation.Add(4)
            Case ButtonFive.Name
                calculation.Add(5)
            Case ButtonSix.Name
                calculation.Add(6)
            Case ButtonSeven.Name
                calculation.Add(7)
            Case ButtonEight.Name
                calculation.Add(8)
            Case ButtonNine.Name
                calculation.Add(9)
            Case ButtonAdd.Name
                calculation.Add("+")
            Case ButtonSubtract.Name
                calculation.Add("-")
            Case ButtonMultiply.Name
                calculation.Add("*")
            Case ButtonDivide.Name
                calculation.Add("/")
            Case ButtonDecimalPoint.Name
                calculation.Add(".")
            Case ButtonCE.Name
                If calculation.Count > 0 Then
                    calculation.RemoveAt(calculation.Count - 1)
                End If
            Case MysteryButton.Name
                Dim randInt As New Random
                calculation.Add(randInt.Next(1, 999))
            Case ButtonC.Name
                ClearAll(True)
        End Select
        UpdateDisplay(calculation)
    End Sub

    Private Function RemoveEmptyValues(ByVal genericList)
        Dim cleanedList As New ArrayList
        For index = 0 To genericList.Length - 1
            Dim num As String = genericList(index)
            If (num = "") = False Then
                cleanedList.Add(num)
            End If
        Next
        Dim finalCleanedList(cleanedList.Count - 1) As String
        For index = 0 To cleanedList.Count - 1
            Dim num As String = cleanedList(index)
            If (num = "") = False Then
                finalCleanedList(index) = num
            End If
        Next
        Return finalCleanedList
    End Function

    Private Sub ResizeLabel()
        Dim labelFont As Font
        Dim labelGraphics As Graphics = ResultLabel.CreateGraphics
        Dim labelSize As SizeF = labelGraphics.MeasureString(ResultLabel.Text, ResultLabel.Font)
        If (ResultLabel.Text.Length = 0) = False Then
            If labelSize.Width > ResultLabel.Width Then
                labelFont = ResultLabel.Font
                ResultLabel.Font = New Font(labelFont.Name, labelFont.SizeInPoints - 1)
            ElseIf (labelSize.Width < ResultLabel.Width) And (labelSize.Height < ResultLabel.Height) Then
                labelFont = ResultLabel.Font
                ResultLabel.Font = New Font(labelFont.Name, labelFont.SizeInPoints + 1)
            End If
        End If
    End Sub

    Private Sub UpdateDisplay(ByVal textToDisplay)
        'ResultLabel.Text = If(TypeOf textToDisplay Is ArrayList, DirectCast(ArrayToString(textToDisplay), String), DirectCast(textToDisplay, String))
        If TypeOf textToDisplay Is ArrayList Then
            ResultLabel.Text = ArrayToString(textToDisplay)
        Else
            ResultLabel.Text = textToDisplay
        End If
        ResizeLabel()
    End Sub

End Class
