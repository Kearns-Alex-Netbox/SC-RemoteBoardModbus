Public Class Main
	#Const DEBUGLOG = 1
	#Const DEBUG = 1

	Dim mbClass As New ModbusXfce()

	Dim valueTBs As Label()

	Private Sub Main_Load() Handles MyBase.Load
		' get our current list of COM ports attached to the computer
		For Each sp As String In My.Computer.Ports.SerialPortNames
			COMPortList.Items.Add(sp.Substring(LEN_STR_COM))
		Next

		' If we have at least one, select the first one in the list
		If 0 < COMPortList.Items.Count Then
			COMPortList.SelectedIndex = 0
		End If
		BaudRate.SelectedIndex = 0

		' initilize our array
#Region "Labels"
		valueTBs = New Label() {
			MB0000,
			MB0001,
			MB0002,
			MB0003,
			MB0004,
			MB0005,
			MB0006,
			MB0007,
			MB0008,
			MB0009	
		}
#End Region

	End Sub

	Private Sub B_Connect_Click() Handles B_Connect.Click
		' setup our outputlog to help debuging
#If DEBUGLOG = 1
		mbClass.EnableLogging("c:\temp\RemoteBoardModbusClient.txt")
        mbClass.EnableMessageLogging()
#End If

		' setup the byte swapping that will be used in communication
		mbClass.SetSwapFlag(ModbusXfce.SwapByteFlag Or ModbusXfce.SwapWordFlag)

		Dim ComPort As Integer = COMPortList.SelectedItem.ToString

		' try to open the connection to the COM Port
		If (mbClass.OpenRTU(SlaveId.Value, ComPort, BaudRate.SelectedItem.ToString, DATABITS, PARITY_NONE, STOPBITS) = False) Then
			MsgBox("Could not open Modbus/RTU port: " & ComPort & vbNewLine & "Slave Id: " & SlaveId.Value)
			ConnectionStatus.Text = "Not Connected"
            Return
        End If

		' lock our setup interfaces and update our status
		ConnectionStatus.Text = "RTU Connected"
		B_Disconnect.Enabled  = True
		B_Connect.Enabled     = False
		COMPortList.Enabled   = False
        SlaveId.Enabled       = False
		BaudRate.Enabled      = False

		GetInitialvalues()
	End Sub

	Private Sub B_Disconnect_Click() Handles B_Disconnect.Click
		' AJK: Need to close any running threads first

		' close our connection
		mbClass.CloseHandle()

		' unlock our setup interfaces and update our status
        ConnectionStatus.Text = "Not Connected"
		B_Disconnect.Enabled  = False
		B_Connect.Enabled     = True
		COMPortList.Enabled   = True
        SlaveId.Enabled       = True
		BaudRate.Enabled      = True
	End Sub

	Private Sub B_Close_Click() Handles B_Close.Click
		Application.Exit
	End Sub

	Private Sub GetInitialvalues()
		' create an error string to show if any issues occured.
		dim errorMessage As String = ""

		' start at the beginning
		dim currentState As MB_ADDRESS = MB_ADDRESS.MB_PRODUCT_CODE_REG

		'While currentState <> MB_ADDRESS.MB_END_OF_ADDRESSES
		While currentState <= MB_ADDRESS.MB_RUNMODE_REG

			dim currentIndex As Integer = Array.IndexOf([Enum].GetValues(currentState.GetType()),currentState)
			dim currentName As String = [Enum].GetName(GetType(MB_ADDRESS), currentState)
			dim str_format As String = "X4"
			dim value As UInteger
			dim retval = 0

#If DEBUG = 1 Then
			' check the enum's name to see if we are a word or a double word
			If currentName.Contains("16") Then
				retval = mbClass.GetWord(currentState, value)

			ElseIf currentName.Contains("32") Then
				str_format = "X8"
				retval = mbClass.GetDword(currentState, value)

			Else
				retval = mbClass.GetWord(currentState, value)

			End If
#Else
			retval = mbClass.GetWord(currentState, value)
#End If

			' check to see that we were successful in getting the value
			If retval <> 0 Then
				errorMessage &= "Error on " & currentName & ":" & vbNewLine & GetMBErrorMessage(retval)
				valueTBs(currentIndex).Text = "ERROR"

			ElseIf valueTBs.Count < currentIndex
				valueTBs(currentIndex).Text = "Index is out of range for our value labels"

			Else
				valueTBs(currentIndex).Text = "0x" & value.ToString(str_format)
			End If

			' bump our state up
			NextEnum_MB_ADDRESS(currentState)
		End While

		' check to see if we have an error message to display
		If errorMessage.Length <> 0 Then
			dim answer As Integer = MessageBox.Show("The following errors occured. Do you still want to continue?" & vbNewLine & 
													errorMessage, "Continue?", MessageBoxButtons.YesNo)

			' check to see if the user wants to quit
			If answer = DialogResult.No
				' close the connection
				B_Disconnect_Click()
				Return
			End If
		End If

	End Sub

	Private Sub MB0004_TextChanged() Handles MB0004.TextChanged
		dim value As Integer = 0

		' try to convert the value to a number
		try
			' The value will already be in a hex form and contain "0x" in the beginning
			value = CINT(MB0004.Text.Replace("0x", "&H"))
		Catch ex As Exception
			' set all of our button to red and 'E's
			Dip0.Text = "E"
			Dip1.Text = "E"
			Dip2.Text = "E"
			Dip3.Text = "E"

			Dip0.BackColor = Color.LightGray
			Dip1.BackColor = Color.LightGray
			Dip2.BackColor = Color.LightGray
			Dip3.BackColor = Color.LightGray

			Return
		End Try

		' now do bitwise operations to figure out which one is on/off
		if ((value And (2^0)) <> BIT_OFF) Then
			Dip0.BackColor = Color.Green
			Dip0.Text = BIT_ON
		else
			Dip0.BackColor = Color.Red
			Dip0.Text = BIT_OFF
		End If
		
		if ((value And (2^1)) <> BIT_OFF) Then
			Dip1.BackColor = Color.Green
			Dip1.Text = BIT_ON
		else
			Dip1.BackColor = Color.Red
			Dip1.Text = BIT_OFF
		End If
		
		if ((value And (2^2)) <> BIT_OFF) Then
			Dip2.BackColor = Color.Green
			Dip2.Text = BIT_ON
		else
			Dip2.BackColor = Color.Red
			Dip2.Text = BIT_OFF
		End If
		
		if ((value And (2^3)) <> BIT_OFF) Then
			Dip3.BackColor = Color.Green
			Dip3.Text = BIT_ON
		else
			Dip3.BackColor = Color.Red
			Dip3.Text = BIT_OFF
		End If
	End Sub

End Class
