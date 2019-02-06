'_____________________ Schedule a task on WIndows Task Scheduler _____________________________________

'Requires TaskScheduler namespace to be imported in UiPath
'Download the package from https://www.nuget.org/packages/TaskScheduler/2.8.7
'Manage Packages > Settings > User defined Package source 
'Add location of downloaded path
'Click on all packages > Search for TaskScheduler and install the package
'Import the Microsoft.Win32.TaskScheduler namespace

Using tService As New TaskService()
	
	Dim tDefinition As TaskDefinition = tService.NewTask
    
		'Description of the task when it appears on the task scheduler
	tDefinition.RegistrationInfo.Description = "Example Task"
 
    Dim tTrigger As New WeeklyTrigger()
	
		'Start datetime in the format (year, month, day, hour, minute, second)
	Dim startTime As New System.DateTime(2019, 1, 15, 11, 15, 0)  
	
		'End datetime in the format (year, month, day, hour, minute, second)
	Dim endTime As New System.DateTime(2019, 1, 15, 11, 30, 0)  
	tTrigger.StartBoundary = startTime
	tTrigger.EndBoundary = endTime
         
		'Trigger on Tuesdays and Sundays
	tTrigger.DaysOfWeek = DaysOfTheWeek.Tuesday Or DaysOfTheWeek.Sunday
    tDefinition.Triggers.Add(tTrigger)
 
		'Action to be executed by the scheduler - here it opens a notepad file at a given location
    tDefinition.Actions.Add(New ExecAction("notepad.exe", "C:\Users\ndh00145\Desktop\trial.txt"))
	
		'Name of the task 
	tService.RootFolder.RegisterTaskDefinition("TrialTask", tDefinition)
 
End Using

'_________________________________________________________________________________________________


'_____________________ Send emails with custom options set _____________________________________

'Useful for adding custom options to outlook emails such as Voting buttons, Read Reciepts etc.
'Requires Microsoft.Office.Interop.Outlook namespace to be imported in UiPath

Dim oApp As New Microsoft.Office.Interop.Outlook.Application
Dim mitem As Microsoft.Office.Interop.Outlook.MailItem

	'Path of the .msg file
	'In Outlook, after composing a blank mail template (with options such as voting buttons set) click File > Save As > Outlook Message Format
mitem = CType(oApp.CreateItemFromTemplate("C:\Users\ndh00145\Desktop\Untitled.msg"),Microsoft.Office.Interop.Outlook.MailItem)

	'To Address / Recipient
mitem.To = "admin-admin@nissanmotor.com"

	'Mail Subject
mitem.Subject = "Test Mail"

	'Mail Body
	mitem.Body = "Hello there"
	
	'Send the email
	mitem.Send()

'_________________________________________________________________________________________________


'_____________________ Convert HTML pages to Excel Worksheet _____________________________________

'Requires Microsoft.Office.Interop.Excel namespace to be imported in UiPath

Dim xlsApp As Microsoft.Office.Interop.Excel.Application = Nothing
Dim xlsWorkBooks As Microsoft.Office.Interop.Excel.Workbooks = Nothing
Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook = Nothing

xlsApp = New Microsoft.Office.Interop.Excel.Application

	'Make Excel visible
xlsApp.Visible = False 

	'Disable alerts such as Save confirmation dialog box
xlsApp.DisplayAlerts = False 
xlsWorkBooks = xlsApp.Workbooks

	'Path to the HTML file
xlsWB = xlsWorkbooks.Open("C:\Users\ndh00145\Documents\My Received Files\email.html") 

	'Path where the Excel workbook is to be stored 
	'Other extensions such as .xls, .xlsm can also be used
xlsWB.SaveAs("C:\Users\ndh00145\Desktop\Mail.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,  System.Reflection.Missing.Value, System.Reflection.Missing.Value, False, False, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, True, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
xlsWB.Close

'_________________________________________________________________________________________________


'_________________________ Invoke a macro on an Excel workbook ___________________________________

'Requires Microsoft.Office.Interop.Excel namespace to be imported in UiPath

Dim xlsApp As Microsoft.Office.Interop.Excel.Application = Nothing
Dim xlsWorkBooks As Microsoft.Office.Interop.Excel.Workbooks = Nothing
Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook = Nothing

xlsApp = New Microsoft.Office.Interop.Excel.Application

	'Make Excel visible
xlsApp.Visible = False 

	'Disable alerts such as Save confirmation dialog box
xlsApp.DisplayAlerts = False 
xlsWorkBooks = xlsApp.Workbooks

	'Path to the Excel Workbook
xlsWB = xlsWorkbooks.Open("C:\Users\ndh00145\Desktop\Mail.xlsx")

	'Path of the macro
	'From the Visual Basic Editor in Excel, Word etc. right click a module and select Export. The file can be saved with the ".bas" extension
xlsApp.VBE.ActiveVBProject.VBComponents.Import("C:\Users\ndh00145\Desktop\FormatterCode.bas")

	'Formatter() is the name of the Sub declared in VBE, i.e., it is the name of the macro
xlsApp.Run("Formatter")

	'Save changes to file
xlsWB.Save

'_________________________________________________________________________________________________


'___________________________________________ HTTP Requests _______________________________________

Dim request As HttpWebRequest  
Dim response As HttpWebResponse = Nothing  
Dim reader As StreamReader  
  
Try  
		' Create the web request  
    request = DirectCast(WebRequest.Create("https://cdn.pixabay.com/photo/2015/05/15/14/38/computer-768608_1280.jpg"), HttpWebRequest)  
  
		' Get response  
    response = DirectCast(request.GetResponse(), HttpWebResponse)  
  
		' Get the response stream (image) into a reader and save to a location 
	Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(response.GetResponseStream())
	img.Save("C:\Users\ndh00145\Desktop\trial.jpeg", System.Drawing.Imaging.ImageFormat.jpeg)

Catch ex As Exception
        ' Show the exception's message.
    MsgBox.Show(ex.Message)    
		
Finally  
    If Not response Is Nothing 
		Then response.Close() 
	End If
End Try  

'_________________________________________________________________________________________________


'___________________________________ Custom Forms with VB .NET _____________________________________

'Requires the System.Windows.Forms namespace to be imported in UiPath

Dim frm As Form = New Form
 
	'Form Height
frm.Height = 500

	'Form Width	
frm.Width = 800

	'Allow scroll bars to appear automatically if form content is large
frm.AutoScroll = True

	'Label for displaying text
Dim Label1 As New Label

	'Position of Label relative to top left corner of the form
Label1.Location = New System.Drawing.Point(500, 50)

	'Label Height
Label1.Width = 300

	'Label Width
Label1.Height = 25

	'Text to be displayed
Label1.Text = "Hello"

	'Label background colour
Label1.BackColor = Color.Aqua
Label1.AutoSize = False

	'Center the text
Label1.TextAlign  = ContentAlignment.MiddleCenter

	'Add the label to the form
frm.Controls.Add(Label1)

	'Triggered upon mouse movement
AddHandler frm.MouseMove, Sub(sender As Object,ByVal e As System.Windows.Forms.MouseEventArgs)
   
		'Display mouse coordinates when inside the form
   Label1.Text = "Mouse X : " & e.X &" | Mouse Y : " & e.Y & " | Mouse Button : " & e.Button
End Sub

	'Title of the custom form
frm.Text = "Custom Form"

	'Load an icon as a Bitmap from file
Dim b As Bitmap 
b = CType(System.Drawing.Image.FromFile("C:\Users\ndh00145\Downloads\iconfinder_katana-simple_479473.png"), Bitmap)
Dim pIcon As New IntPtr 
pIcon = b.GetHicon()
Dim i As Icon 
i = Icon.FromHandle(pIcon)

	'set the image as the form's icon
frm.Icon = i
i.Dispose()
   
   'Triggered whe close button is clicked
   'Prevents form from closing when the invoke code activity finishes 
AddHandler frm.FormClosing, Sub(sender As Object,ByVal e As System.Windows.Forms.FormClosingEventArgs)
   frm.Close()
End Sub

	'Create a new combobox (dropdown list)
Dim ComboBox1 As New ComboBox

	'location of the combobox
ComboBox1.Location = New System.Drawing.Point(12, 12)

	'Name of Combobox
ComboBox1.Name = "ComboBox1"

	'Size of ComboBox
ComboBox1.Size = New System.Drawing.Size(245, 25)

	'Background Color
ComboBox1.BackColor = System.Drawing.Color.Orange

	'Foreground Color
ComboBox1.ForeColor = System.Drawing.Color.Black

	'Height to which dropdown should extend downwards
ComboBox1.DropDownHeight = 100

	'Width of the dropdown list
ComboBox1.DropDownWidth = 300

	'Values for the list
ComboBox1.Items.Add("Mahesh Chand")
ComboBox1.Items.Add("Mike Gold")
ComboBox1.Items.Add("Praveen Kumar")
ComboBox1.Items.Add("Raj Beniwal")
ComboBox1.DropDownStyle = ComboBoxStyle.DropDown

	'Default placeholder text
ComboBox1.Text = "Select an item"
ComboBox1.SelectionLength = 0
frm.Controls.Add(ComboBox1)

	'Add a picture box to form (Necessary for drawing graphics such as lines)
Dim pictureBox1 As New PictureBox
pictureBox1.Location = New System.Drawing.Point(12, 80)
frm.Controls.Add(pictureBox1)

AddHandler pictureBox1.Paint, Sub(sender As Object,ByVal e As System.Windows.Forms.PaintEventArgs)
  Dim grph As Graphics = frm.CreateGraphics
  With grph
	Dim p As New Pen(Color.Red, 3)

	'Draw a line from point (0,500) to (Width of form, 500), i.e. at a height of 500 with a width equal to length of form
    .DrawLine(p, 0, 500, frm.Width, 500)
  End With
End Sub

	'Create a button 
Dim btn As Button = New Button
btn.Location = New System.Drawing.Point(300, 50)
btn.Text = "Exit"
frm.Controls.Add(btn)

AddHandler btn.Click, Sub(sender As Object,ByVal e As System.EventArgs)
   frm.close()

End Sub

	'Create a menu bar
Dim mnuBar As New MainMenu()
      
	'defining the menu item named "File" for the main menu bar
Dim myMenuItemFile As New MenuItem("&File")
mnuBar.MenuItems.Add(myMenuItemFile)
    
	'defining the sub menu item "New" within the File menu
Dim myMenuItemNew As New MenuItem("&New")
myMenuItemFile.MenuItems.Add(myMenuItemNew)

	'Add menubar to form
frm.Menu = mnuBar
   
   
	'Display the form
frm.ShowDialog()
 
'____________________________________________________________________________________________________


