Public Module Common

	Public Const PARITY_EVEN							As Byte    = 101   ' even
    Public Const PARITY_ODD								As Byte    = 111   ' odd
    Public Const PARITY_NONE							As Byte    = 110   ' none

	Public Const BAUDRATE								As Integer = 225600'115200
	Public Const DATABITS								As Integer = 8
	Public Const STOPBITS								As Integer = 1

	Public Const STR_COM								As String  = "COM"
	Public       LEN_STR_COM							As Integer = STR_COM.Length

	Public Const BIT_OFF								As Integer = 0
	Public Const BIT_ON									As Integer = 1

	Public Const DIP_SWITCH_LOADER_INDEX				As Integer = 3
	Public Const DIP_SWITCH_COMM_PACKET_COUNT_INDEX		As Integer = 2


	Public Const REMOTEIO_RUNMODE_LOADER				As Integer = 1		' we are running in Loader mode
	Public Const REMOTEIO_RUNMODE_APPLICATION			As Integer = 2		' we are running in Application mode

	' Modbus Registers for retrieving data
	Public Enum MB_ADDRESS
		MB_PRODUCT_CODE_REG					= 0		' send product code for this board (one of MODULE_TYPE_xxx)
		MB_PRODUCT_VERSION_REG				 '1		' send software version running (HHLL)
		MB_DIGITAL_COUNT_REG				 '2		' send number of digital inputs/digital outputs (IIOO)
		MB_ANALOG_COUNT_REG					 '3		' send number of analog inputs/analog outputs (IIOO)
		MB_DIPSWITCH_REG					 '4		' send DIP switch settings and option bits (OOxxxxxxxxxx3210) OO are 2 option bits, 3210 are DIP switch #'s
		MB_CPU_FREQUENCY_REG				 '5		' running CPU frequency in MHz (32 bits)
		MB_CPU_FREQUENCY_HIGH_REG			 '6		' running CPU frequency in MHz (32 bits)
		MB_TIMER_PRESCALER_REG				 '7		' timer prescaler reg (32 bits)
		MB_TIMER_PRESCALER_HIGH_REG			 '8		' timer prescaler reg (32 bits)
		MB_RUNMODE_REG						 '9		' is it running application or loader (1=loader, 2=application)
		
		MB_DIGITAL_INPUT1_REG				= 109	' tank
		MB_DIGITAL_INPUT2_REG				 '110
		MB_ANALOG_INPUT1_RAW_REG			 '111
		MB_ANALOG_INPUT2_RAW_REG			 '112
		MB_ANALOG_INPUT3_RAW_REG			 '113
		MB_ANALOG_INPUT4_RAW_REG			 '114
		MB_DIGITAL_FREQ1_CURRENT_LOW_REG	 '115
		MB_DIGITAL_FREQ1_CURRENT_HIGH_REG	 '117
		MB_DIGITAL_OUTPUT1_READ_REG			 '119
		MB_DIGITAL_OUTPUT2_READ_REG			 '120
												
		
		MB_DIGITAL_OUTPUT1_REG				= 200	
		MB_DIGITAL_OUTPUT2_REG				 '201
		MB_ANALOG_OUTPUT1_RAW_REG			 '202
		MB_ANALOG_OUTPUT2_RAW_REG			 '203
												
		MB_COMMAND_REG						= 300	' command register for calibration
												
		MB_ANALOG_INPUT1_CAL_BIAS_REG		= 400	
		MB_ANALOG_INPUT2_CAL_BIAS_REG		 '401
		MB_ANALOG_INPUT3_CAL_BIAS_REG		 '402
		MB_ANALOG_INPUT4_CAL_BIAS_REG		 '403
		MB_ANALOG_INPUT1_CAL_GAIN_REG		 '404
		MB_ANALOG_INPUT2_CAL_GAIN_REG		 '405
		MB_ANALOG_INPUT3_CAL_GAIN_REG		 '406
		MB_ANALOG_INPUT4_CAL_GAIN_REG		 '407
												
		MB_ANALOG_OUTPUT1_CAL_BIAS_REG		= 500	
		MB_ANALOG_OUTPUT1_CAL_GAIN_REG		 '501
		MB_ANALOG_OUTPUT2_CAL_BIAS_REG		 '502
		MB_ANALOG_OUTPUT2_CAL_GAIN_REG		 '503
		
												
		MB_DIGITALV2_INPUT1_REG				= 1009
		MB_DIGITALV2_INPUT2_REG				 '1010
		MB_ANALOGV2_INPUT1_RAW_REG			 '1011
		MB_ANALOGV2_INPUT2_RAW_REG			 '1012
		MB_ANALOGV2_INPUT3_RAW_REG			 '1013
		MB_ANALOGV2_INPUT4_RAW_REG			 '1014
		MB_ANALOGV2_INPUT5_RAW_REG			 '1015
		MB_ANALOGV2_INPUT6_RAW_REG			 '1016
		MB_ANALOGV2_INPUT7_RAW_REG			 '1017
		MB_ANALOGV2_INPUT8_RAW_REG			 '1018
		
												
		MB_SLAVE_STAT_LAST_ANALOG_COMP_TIME	= 3000
		MB_SLAVE_STAT_MAX_IRQ_TIME			 '3002
		MB_SLAVE_ADC_FAIL_COUNT				 '3004
		MB_SLAVE_ADC_VREF_INT				 '3005
		MB_SLAVE_ADC_VREF_CALIBRATE			 '3006
												
		MB_SLAVE_STAT_UART_ERRORS			= 3008	
		MB_SLAVE_STAT_PACKETLEN_ERRORS		 '3010
		MB_SLAVE_STAT_UART_STATUS			 '3012
		MB_SLAVE_STAT_TIMEOUT_ERRORS		 '3014
		MB_SLAVE_STAT_INTERCHTIMEOUT_ERRORS	 '3016
		MB_SLAVE_STAT_PACKET_COUNT			 '3018

		' needed to signal the absolute end of the enums
		MB_END_OF_ADDRESSES                 
	End Enum



	Public sub NextEnum_MB_ADDRESS(ByRef currentState As Object)
		' make sure that the passed in state is not the last of its kind
		if (currentState = [Enum].GetValues(GetType(MB_ADDRESS)).Cast(Of MB_ADDRESS).Last())
			Return
		End If

		' increase our state to the next increment
		currentState += 1
		
		' only loop if we are not a defined enum state
		' the enum state should have a END value to prevent from continual looping.
		while [Enum].IsDefined(GetType(MB_ADDRESS), currentState) = false 
			currentState += 1
		End While
	End sub

	Public Function GetMBErrorMessage(ByVal errorCode As Integer) As String
		Dim errorMessage As String = ""
		Select Case errorCode
			Case 0
				errorMessage = "Successful. No Error."
			Case 1
				errorMessage = "Error Message 1."
			Case Else
				errorMessage = "Message has not been handled yet."
		End Select

		Return errorMessage
	End Function

End Module