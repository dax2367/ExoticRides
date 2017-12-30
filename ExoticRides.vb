Option Strict On

Public Class ExoticRides
    '================================================================================================================
    ' Author:	Holly Eaton
    ' 
    ' Program:  Exotic Rides
    ' 
    ' Dev Env:	Visual Studio
    ' 
    ' Description:
    '   Purpose:    Project that will determine:
    '                   Rental charges for a single customer.               (Customer Rental Charge)
    '                   The number of people who have rented cars.		    (Total Rentals)
    '                   The total number of rental hours for each customer. (Total Hours)
    '                   The total of all rentals fees.			            (Total Charges)
    '                   The average rental charge for each customer.	    (Average Charges)
    '  
    '   Input:      Customer Name, Hours rented, AAA-membership. 
    '
    '   Process:    Calculate the following:
    '                   The individual customer rental charge.
    '                   Increment the total number of rentals processed.
    '                   Add the number of hours for the current customer to the total number of hours recorded so far.
    '                   Add the charges for the current customer to the total amount of rental charges recorded so far.
    '                   Calculate the average rental charge per customer.
    '  
    '   Output:     Textual information for the user inside the labels(totals) and textboxes(name, hours rented).
    '                   Format (as currency) and display the individual rental customer charge, the total recorded
    '                   charges, and the average charge inside of appropriate labels.
    '                   Format (as numbers) and display the total number of rentals and the total number of hours
    '                   recorded so far in the appropriate labels. 
    ' 
    '==================================================================================================================
    ' 	Declared Constants:
    '	dblREG_RENTAL_RATE
    '	dblGOLD _RATE
    '	dblPLATINUM _RATE
    '	dblTITANIUM _RATE
    '	dblAAA_ DISCOUNT
    '	dblGOLD_CUTOFF
    '	dblPLATINUM_CUTOFF
    '	dblTITANIUM_CUTOFF
    '	dblAAA_CUTOFF
    '
    '==================================================================================================================
    '	Variables for user entered data:
    '	strCustName
    '	dblCustHoursThisRental
    '	Note: (dim and cast both directly to double then no type casting necessary in calculations)
    '	Example: Dim dblCustHoursThisRental As Double = Cdbl(txtCustHoursThisRental.Text)
    '
    '==================================================================================================================
    '	Variables for calculated values:
    '	dblCustHoursThisRental
    '	dblCustRentalCharge
    '
    '   *************************************************
    '	***Class level variables for calculated values***
    '	***         dblTotalRentals                   ***
    '	***         dblTotalHours                     ***
    '	***         dblTotalCharges                   ***
    '	***         dblAveOfCharges                   ***
    '	*************************************************                
    '
    '===================================================================================================================
    '   Calculations in pseudocode:
    '	Set option strict on at top
    '   CustRentalCharge= (CustHoursThisRental * (Appropriate RentalRate)) - (AAAdiscount if applicable)
    '   TotalRentals += 1                   (same as TotalRentals = TotalRentals + 1)
    '   TotalHours += CustHoursThisRental   (same as TotalHours = TotalHours + CustHoursThisRental)
    '   TotalCharges += CustRentalCharge    (same as TotalCharges = TotalCharges + CustHoursThisRental)
    '   AveOfCharges = TotalCharges / TotalRentals
    '
    '====================================================================================================================
    '====================================================================================================================
    '====================================================================================================================

    'Declared Constants:
    Const dblREG_RENTAL_RATE As Double = 99.99
    Const dblGOLD_RATE As Double = 89.49105
    Const dblPLATINUM_RATE As Double = 84.9915
    Const dblTITANIUM_RATE As Double = 79.992
    Const dblAAA_DISCOUNT As Double = 149.95
    Const dblGOLD_CUTOFF As Double = 3
    Const dblPLATINUM_CUTOFF As Double = 10
    Const dblTITANIUM_CUTOFF As Double = 24
    Const dblAAA_CUTOFF As Double = 15

    'Create module variables to keep track of accumulating totals
    Private dblTotalRentals As Double
    Private dblTotalHours As Double
    Private dblTotalCharges As Double
    Private dblAveOfCharges As Double

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Dim dblCustRentalCharge As Double
        Dim strCustName As String
        If txtCustName.Text <> String.Empty Then
            strCustName = txtCustName.Text
            Try
                Dim dblCustHoursThisRental As Double = CDbl(txtCustHoursThisRental.Text)

                'Calculations for the Customer Rental Charge, Total Rentals, Total Hours, Total Charges, and Average Charges.
                'CustRentalCharge
                If dblCustHoursThisRental <= dblGOLD_CUTOFF Then
                    dblCustRentalCharge = dblCustHoursThisRental * dblREG_RENTAL_RATE
                ElseIf dblCustHoursThisRental <= dblPLATINUM_CUTOFF Then
                    dblCustRentalCharge = dblCustHoursThisRental * dblGOLD_RATE
                ElseIf dblCustHoursThisRental <= dblTITANIUM_CUTOFF Then
                    dblCustRentalCharge = dblCustHoursThisRental * dblPLATINUM_RATE
                Else : dblCustRentalCharge = dblCustHoursThisRental * dblTITANIUM_RATE
                End If

                If chkAAAMembr.Checked = True Then
                    If dblCustHoursThisRental > dblAAA_CUTOFF Then
                        dblCustRentalCharge -= dblAAA_DISCOUNT
                    End If
                End If

                'TotalRentals
                dblTotalRentals += 1

                'TotalHours
                dblTotalHours += dblCustHoursThisRental

                'TotalCharges
                dblTotalCharges += dblCustRentalCharge

                'AveOfCharges
                dblAveOfCharges = dblTotalCharges / dblTotalRentals

                'format and output the results
                lblCustomerRentalCharge.Text = dblCustRentalCharge.ToString("c")
                lblTotalRentals.Text = dblTotalRentals.ToString("n0")
                lblTotalHours.Text = dblTotalHours.ToString("n0")
                lblTotalCharges.Text = dblTotalCharges.ToString("c")
                lblAveOfCharges.Text = dblAveOfCharges.ToString("c")

                'Send focus to the Clear button
                btnClear.Focus()
            Catch ex As Exception
                'What to do if user data entered into txtCustHoursThisRental is invalid and cannot be cast to dbl.
                'Tell user what to enter
                MessageBox.Show("Please enter the number of hours the customer rented the vehicle using only numerical characters.")
                'Reset input area
                txtCustHoursThisRental.Text = String.Empty
                'Put insertion point inside of Hours Rented textbox.
                txtCustHoursThisRental.Focus()
            End Try
        Else
            MessageBox.Show("Please enter the customers name")
            txtCustName.Focus()
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear the textboxes and labels
        txtCustName.Text = String.Empty
        txtCustHoursThisRental.Text = String.Empty
        lblCustomerRentalCharge.Text = String.Empty
        chkAAAMembr.Checked = False

        'Give the focus to txtNumLgCookies
        txtCustName.Focus()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Close the form
        Me.Close()
    End Sub
End Class
