# Import-CSV-Wizard
Powershell GUI that'll import users from a CSV file into Active Directory


[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();

Add-Type -assembly System.Windows.Forms
$ExcelPicture = (get-item excel.jpg)
$BackGroundPic = (get-item bg10.jpg)

$Main_Form = New-Object System.Windows.Forms.Form
$main_form.Text ='Import CSV Wizard'
$main_form.Width = 600
$main_form.Height = 600
$main_form.StartPosition = "CenterScreen"
$Main_Form.BackColor = "Gray"
$Main_Form.BringToFront()


$excel = [System.Drawing.Image]::FromFile($ExcelPicture)
$ExcelLogo = New-Object System.Windows.Forms.PictureBox
$ExcelLogo.location = New-Object System.Drawing.Size (100,100)
$ExcelLogo.Width = 125
$ExcelLogo.Height = 110
$ExcelLogo.left = 20
$ExcelLogo.top = 20
$ExcelLogo.sizemode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$ExcelLogo.Image = $excel
$ExcelLogo.BringToFront()
$main_form.Controls.Add($ExcelLogo)
$main_Form.Add_Shown({$Main_Form.Activate()})


$Label = New-Object System.Windows.Forms.Textbox
$Label.Text = "Greetings, This wizard will instruct you on how to quickly create users in Active Directory"
$Label.Font = New-Object System.Drawing.Font("Times New Roman",14)
$Label.Size = New-Object System.Drawing.Size(200,100)
$Label.Location = New-Object System.Drawing.Point(200,150)
$Label.Multiline = $true
$Label.Enabled = $false
$main_form.Controls.Add($Label)


Function Button_Click()
{
    $Second_Page = New-Object System.Windows.Forms.Form
    $Second_Page.Text = 'Import CSV Wizard'
    $Second_Page.Width = 600
    $Second_Page.Height = 600
    $Second_Page.StartPosition = 'CenterScreen'

    $Second_Page_Text = New-Object System.Windows.Forms.TextBox
    $Second_Page_Text.Text = "Please click one of the master CSV files below, and edit as necessary before proceding to the next screen. After editing, save the file in an easily accessible location."
    $Second_Page_Text.Multiline = $true
    $Second_Page_Text.Location = New-Object System.Drawing.Point (200,100)
    $Second_Page_Text.Size = New-Object System.Drawing.Size (245,140)
    $Second_Page_Text.Font = New-Object System.Drawing.Font ("Times New Roman",14)
    $Second_Page_Text.ReadOnly = $true

    $NotePad_CSV = New-Object System.Windows.Forms.Button
    $NotePad_CSV.Location = New-Object System.Drawing.Point (125,430)
    $NotePad_CSV.Text = "NotePad CSV Format."
    $NotePad_CSV.Height = 100
    $NotePad_CSV.Width = 90

    $Excel_CSV = New-Object System.Windows.Forms.Button
    $Excel_CSV.Location = New-Object System.Drawing.Point (350,430)
    $Excel_CSV.Text = "Master CSV file."
    $Excel_CSV.Height = 100
    $Excel_CSV.Width = 90

    Function Button_Click4()
    {
    $Third_Page = New-Object System.Windows.Forms.Form
    $Third_Page.text = "Import CSV Wizard"
    $Third_Page.Width = 600
    $Third_Page.Height = 600
    $Third_Page.StartPosition = "CenterScreen"

    Function Button_Click5()
    {
    $Open_FD = New-Object System.Windows.Forms.OpenFileDialog
    $Open_FD.Title = "Select .CSV File"
    $Open_FD.InitialDirectory = "C:\"
    $Open_FD.Filter = "CSV Files (*.csv)|*.csv"
    $Open_FD.ShowDialog() | Out-Null
    $Open_FD.filename

    }
    
    $ThirdPageText = New-Object System.Windows.Forms.TextBox
    $ThirdPageText.Text = "Click the Browse button below, and select the previously saved CSV File. Then, press Import."
    $ThirdPageText.Multiline = $True
    $ThirdPageText.Location = New-Object System.Drawing.Point (200,100)
    $ThirdPageText.Size = New-Object System.Drawing.Size (245,100)
    $ThirdPageText.Font = New-Object System.Drawing.Font ("Times New Roman",14)
    $ThirdPageText.ReadOnly = $true


    $OutputBox = New-Object System.Windows.Forms.TextBox
    $OutputBox.Size = New-Object System.Drawing.Size (300,35)
    $OutputBox.Location = New-Object System.Drawing.Point (100,300)
    $Outputbox.Multiline = $true


    $BrowseBtn = New-Object System.Windows.Forms.Button
    $BrowseBtn.Location = New-Object System.Drawing.Point(410,290)
    $BrowseBtn.Text = "Browse"
    $BrowseBtn.Height= 50
    $BrowseBtn.Width=100
    $BrowseBtn.Add_Click({$OutputBox.Text = Button_Click5 -InitialDirectory "C:\"})


    Function Button_Click6()
    {
    Import-CSV $OutputBox.text | New-ADUser
    $Third_Page.Close()
    }

    $ImportAD = New-Object System.Windows.Forms.Button
    $ImportAD.Location = New-Object System.Drawing.Point (250,400)
    $ImportAD.Text = "Import"
    $ImportAD.Height = 50
    $ImportAD.Width = 100
    $ImportAD.Add_Click({Button_Click6})

    $Third_page.Controls.Add($ThirdPageText)
    $Third_Page.Controls.Add($ImportAD)
    $Third_Page.Controls.Add($OutputBox)
    $Third_Page.Controls.Add($BrowseBtn)
    $Third_Page.Controls.Add($ExcelLogo)
    $Third_Page.Controls.Add($BackGround)
    $Third_Page.ShowDialog()
    $Second_Page.Close()
    }


    Function Button_Click2()
    {
    Invoke-Item NotePad.csv
    }


    Function Button_Click3()
    {
    Invoke-Item MasterCSV.csv
    }

    $NextPageButton2 = New-Object System.Windows.Forms.Button
    $NextPageButton2.Location = New-Object System.Drawing.Point (500,475)
    $NextPageButton2.Text = "Next"
    $NextPageButton2.Height = 30
    $NextPageButton2.Width = 60
    $NextPageButton2.Add_Click({Button_Click4})


    $NotePad_CSV.Add_Click({Button_Click2})
    $Excel_CSV.Add_Click({Button_Click3})
    $Second_Page.Controls.Add($Second_Page_Text)
    $Second_Page.Controls.Add($ExcelLogo)
    $Second_Page.Controls.Add($NotePad_CSV)
    $Second_Page.Controls.Add($Excel_CSV)
    $Second_Page.Controls.Add($NextPageButton2)
    $Second_Page.Controls.Add($BackGround)
    $Second_Page.ShowDialog()
    $Main_Form.Close()
    }

$NextPageButton = New-Object System.Windows.Forms.Button
$NextPageButton.Location = New-Object System.Drawing.Point (500,475)
$NextPageButton.Text = "Next"
$NextPageButton.Height = 30
$NextPageButton.Width = 60
$NextPageButton.Add_Click({Button_Click})
$Main_Form.Controls.Add($NextPageButton)


$BG = [System.Drawing.Image]::FromFile($BackGroundPic)
$BackGround = New-Object System.Windows.Forms.PictureBox
$BackGround.location = New-Object System.Drawing.Size (0,0)
$BackGround.Width = $BG.Width
$BackGround.Height = $BG.Height
$BackGround.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$BackGround.BackgroundImage = $BG
$BackGround.SendToBack()
$Main_Form.Controls.Add($Background)


$main_form.ShowDialog()
