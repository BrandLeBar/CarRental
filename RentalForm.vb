Option Explicit On
Option Strict On
Option Compare Text
'Brandon Barrera
'RCET 0226
'Spring 2025
'Car Rental
'https://github.com/BrandLeBar/CarRental.git

Imports System.Runtime.Remoting.Messaging

Public Class RentalForm

    ''' <summary>
    ''' Sets all of the initial values and resets everything to the way it was on start-up.
    ''' </summary>
    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
    End Sub



    Function CheckUserInput() As Boolean
        Dim valid As Boolean = True

        If NameTextBox.Text = "" Then
            MsgBox("Please enter your name")
            NameTextBox.Focus()
            valid = False
        ElseIf AddressTextBox.Text = "" Then
            MsgBox("Please enter your street adress")
            AddressTextBox.Focus()
            valid = False
        ElseIf CityTextBox.Text = "" Then
            MsgBox("Please enter a city")
            CityTextBox.Focus()
            valid = False
        ElseIf StateTextBox.Text = "" Then
            MsgBox("Please enter a State")
            StateTextBox.Focus()
            valid = False
        ElseIf IsNumeric(ZipCodeTextBox.Text) = False Then
            MsgBox("Please enter a valid zip code")
            ZipCodeTextBox.Focus()
            valid = False
        ElseIf IsNumeric(BeginOdometerTextBox.Text) = False Then
            MsgBox("Please enter the starting mileage")
            BeginOdometerTextBox.Focus()
            valid = False
        ElseIf IsNumeric(EndOdometerTextBox.Text) = False Then
            MsgBox("Please enter the ending mileage")
            EndOdometerTextBox.Focus()
            valid = False
        ElseIf CInt(BeginOdometerTextBox.Text) > CInt(EndOdometerTextBox.Text) Then
            MsgBox("YOU CANNOT DRIVE NEGATIVE MILES!!!")
            valid = False
        ElseIf IsNumeric(DaysTextBox.Text) = False Then
            MsgBox("Please enter a valid amount of days 1 - 45")
            DaysTextBox.Focus()
            valid = False
        ElseIf CInt(DaysTextBox.Text) < 1 Or CInt(DaysTextBox.Text) > 45 Then
            MsgBox("Please enter a valid amount of days 1 - 45")
            DaysTextBox.Focus()
            valid = False
        End If
        Return valid
    End Function

    Function TotalDiscount() As Decimal
        Dim _totalDiscount As Decimal
        If Seniorcheckbox.Checked And AAAcheckbox.Checked Then
            _totalDiscount = CDec(0.008)
        ElseIf AAAcheckbox.Checked Then
            _totalDiscount = CDec(0.005)
        ElseIf Seniorcheckbox.Checked Then
            _totalDiscount = CDec(0.003)
        Else
            _totalDiscount = 0
        End If

        Return _totalDiscount
    End Function

    Function TotalCharged() As Decimal
        Dim _totalCharged As Decimal

        _totalCharged = TotalMilesCharged() + TotalDaysCharged()

        Return _totalCharged
    End Function

    Function TotalMiles() As Decimal
        Dim _totalMiles As Decimal = CDec(EndOdometerTextBox.Text) - CDec(BeginOdometerTextBox.Text)
        Return _totalMiles
    End Function

    Function TotalMilesCharged() As Decimal
        Dim _totalMilesCharged As Decimal = TotalMiles()

        If _totalMilesCharged < 200 Then
            _totalMilesCharged = 0
        ElseIf _totalMilesCharged > 500 Then
            _totalMilesCharged = _totalMilesCharged - 200
            _totalMilesCharged = _totalMilesCharged * CDec(0.01)
        Else
            _totalMilesCharged = _totalMilesCharged - 200
            _totalMilesCharged = _totalMilesCharged * CDec(0.012)
        End If
        Return _totalMilesCharged
    End Function

    Function TotalDaysCharged() As Decimal
        Dim _totalDaysCharged As Decimal = CInt(DaysTextBox.Text) * 15
        Return _totalDaysCharged
    End Function

    Function CustomerCounter(Optional clear As Boolean = False, Optional referance As Boolean = False) As Integer
        Static _customerCounter As Integer

        If clear = False And referance = False Then
            _customerCounter += 1
        ElseIf clear = True Then
            _customerCounter = 0
        End If

        Return _customerCounter
    End Function

    Function MileageCounter(Optional clear As Boolean = False, Optional referance As Boolean = False) As Decimal
        Static _mileageCounter As Decimal

        If clear = False And referance = False Then
            _mileageCounter += TotalMiles()
        ElseIf clear = True Then
            _mileageCounter = 0
        End If

        Return _mileageCounter
    End Function

    Function TotalChargedCounter(Optional clear As Boolean = False, Optional referance As Boolean = False) As Decimal
        Static _totalChargedCounter As Decimal

        If clear = False And referance = False Then
            _totalChargedCounter += (TotalCharged() - TotalCharged() * TotalDiscount())
        ElseIf clear = True Then
            _totalChargedCounter = 0
        End If

        Return _totalChargedCounter
    End Function



    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem.Click, ExitToolStripMenuItem1.Click
        If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Exit") = MsgBoxResult.Yes Then
            Me.Close()
        Else

        End If

    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetDefaults()
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem.Click, SummaryToolStripMenuItem1.Click
        Dim totalCustomers As Integer = CustomerCounter(, True)
        Dim totalMileage As Decimal = MileageCounter(, True)
        Dim totalCharged As Decimal = TotalChargedCounter(, True)

        MsgBox($"Customers: {totalCustomers} {vbNewLine} Total Mileage: {totalMileage} {vbNewLine} Total Charged: {totalCharged}")
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem.Click, ClearToolStripMenuItem1.Click
        SetDefaults()
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click, FileToolStripMenuItem.Click

        If CheckUserInput() = True Then
            If TotalChargeTextBox.Text = (TotalCharged() - TotalCharged() * TotalDiscount()).ToString("C") Then
                'abosolutly nothin
            Else
                CustomerCounter()
                MileageCounter()
                TotalChargedCounter()

            End If
            TotalMilesTextBox.Text = $"{TotalMiles()} Mi"
            MileageChargeTextBox.Text = TotalMilesCharged().ToString("C")
            DayChargeTextBox.Text = TotalDaysCharged().ToString("C")
            TotalDiscountTextBox.Text = (TotalCharged() * TotalDiscount()).ToString("C")
            TotalChargeTextBox.Text = (TotalCharged() - TotalCharged() * TotalDiscount()).ToString("C")
        End If
    End Sub

    Private Sub KilometersradioButton_CheckedClicked(sender As Object, e As EventArgs) Handles KilometersradioButton.Click
        EndOdometerTextBox.Text = $"{CDec(EndOdometerTextBox.Text) * CDec(0.62)}"
        BeginOdometerTextBox.Text = $"{CDec(BeginOdometerTextBox.Text) * CDec(0.62)}"
    End Sub

    Private Sub MilesradioButton_Clicked(sender As Object, e As EventArgs) Handles MilesradioButton.Click
        EndOdometerTextBox.Text = $"{CDec(EndOdometerTextBox.Text) / CDec(0.62)}"
        BeginOdometerTextBox.Text = $"{CDec(BeginOdometerTextBox.Text) / CDec(0.62)}"
    End Sub

End Class
