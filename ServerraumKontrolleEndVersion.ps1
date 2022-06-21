# Autor: Dominik Costa
# Version: 4.0
# Erstellt: 17. Oktober 2019
# Mutated: 21. Juni 2022

#----------------------------------------------------------------------------------
# Start

PowerShell.exe -windowstyle hidden {

#----------------------------------------------------------------------------------
#GUI Fenster
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$window = New-Object System.Windows.Forms.Form
$window.Width = 700
$window.Height = 300

#----------------------------------------------------------------------------------
#Text Server room 1 aussen
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(10,10)
$Label.Text = "Serverroom 1 outside"
$Label.AutoSize = $True

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Size(10,30)
$Label2.Text = "Temperatur"
$Label2.AutoSize = $True

$Label3 = New-Object System.Windows.Forms.Label
$Label3.Location = New-Object System.Drawing.Size(150,30)
$Label3.Text = "humidity"
$Label3.AutoSize = $True
$window.Controls.Add($Label)
$window.Controls.Add($Label2)
$window.Controls.Add($Label3)

#----------------------------------------------------------------------------------
#Eingabe  Server room 1 aussen

$windowTextBox = New-Object System.Windows.Forms.TextBox
$windowTextBox.Location = New-Object System.Drawing.Size(10,50)
$windowTextBox.Size = New-Object System.Drawing.Size(100,100)
$window.Controls.Add($windowTextBox)

$windowTextBox2 = New-Object System.Windows.Forms.TextBox
$windowTextBox2.Location = New-Object System.Drawing.Size(150,50)
$windowTextBox2.Size = New-Object System.Drawing.Size(100,100)
$window.Controls.Add($windowTextBox2)

#----------------------------------------------------------------------------------
#Text Server room 1 Innen
$Label4 = New-Object System.Windows.Forms.Label
$Label4.Location = New-Object System.Drawing.Size(350,10)
$Label4.Text = "Serverroom 1 inside"
$Label4.AutoSize = $True

$Label5 = New-Object System.Windows.Forms.Label
$Label5.Location = New-Object System.Drawing.Size(350,30)
$Label5.Text = "Temperature"
$Label5.AutoSize = $True

$Label6 = New-Object System.Windows.Forms.Label
$Label6.Location = New-Object System.Drawing.Size(500,30)
$Label6.Text = "Humidity"
$Label6.AutoSize = $True
$window.Controls.Add($Label4)
$window.Controls.Add($Label5)
$window.Controls.Add($Label6)

#----------------------------------------------------------------------------------
# Server room 1 Innen

$windowTextBox3 = New-Object System.Windows.Forms.TextBox
$windowTextBox3.Location = New-Object System.Drawing.Size(350,50)
$windowTextBox3.Size = New-Object System.Drawing.Size(100,100)
$window.Controls.Add($windowTextBox3)

$windowTextBox4 = New-Object System.Windows.Forms.TextBox
$windowTextBox4.Location = New-Object System.Drawing.Size(500,50)
$windowTextBox4.Size = New-Object System.Drawing.Size(100,100)
$window.Controls.Add($windowTextBox4)

#----------------------------------------------------------------------------------
# Text Server room 2
$Label7 = New-Object System.Windows.Forms.Label
$Label7.Location = New-Object System.Drawing.Size(10,80)
$Label7.Text = "Serveroom 2"
$Label7.AutoSize = $True

$Label8 = New-Object System.Windows.Forms.Label
$Label8.Location = New-Object System.Drawing.Size(10,100)
$Label8.Text = "Temperature"
$Label8.AutoSize = $True

$Label9 = New-Object System.Windows.Forms.Label
$Label9.Location = New-Object System.Drawing.Size(150,100)
$Label9.Text = "Humidity"
$Label9.AutoSize = $True
$window.Controls.Add($Label7)
$window.Controls.Add($Label8)
$window.Controls.Add($Label9)

#----------------------------------------------------------------------------------
# Server room 2

$windowTextBox5 = New-Object System.Windows.Forms.TextBox
$windowTextBox5.Location = New-Object System.Drawing.Size(10,120)
$windowTextBox5.Size = New-Object System.Drawing.Size(100,100)
$window.Controls.Add($windowTextBox5)

$windowTextBox6 = New-Object System.Windows.Forms.TextBox
$windowTextBox6.Location = New-Object System.Drawing.Size(150,120)
$windowTextBox6.Size = New-Object System.Drawing.Size(100,100)
$window.Controls.Add($windowTextBox6)

#----------------------------------------------------------------------------------
 #Button
  $windowButton = New-Object System.Windows.Forms.Button
  $windowButton.Location = New-Object System.Drawing.Size(10,170)
  $windowButton.Size = New-Object System.Drawing.Size(50,50)
  $windowButton.Text = "OK"
  $windowButton.Add_Click({
     write-host $windowTextBox.Text
     write-host $windowTextBox2.Text
     write-host $windowTextBox3.Text
     write-host $windowTextBox4.Text
     write-host $windowTextBox5.Text
     write-host $windowTextBox6.Text
     clear
   
#----------------------------------------------------------------------------------
# openExcel

$objexcel=New-Object -ComObject Excel.Application
$workbook=$objexcel.WorkBooks.Open('YourPath\Serverroom.xlsx')
$worksheet=$workbook.WorkSheets.item(1)
$objexcel.Visible=$true
sleep 2

#----------------------------------------------------------------------------------
# Variable
    $zeit = Get-Date -f HH:mm
    $kuerzel = $env:UserName
	$DD = Get-Date -f dd.MM.yyyy
	$Day = Get-Date -f dddd
	$date = "     " + $Day.Remove(2)+ ", " + $DD
	$bemerkung = ""
	
#----------------------------------------------------------------------------------
# check temperature

	if ([float]$windowTextBox.Text -gt 27.0 -or [float]$windowTextBox3.Text -gt 27.0 -or [float]$windowTextBox5.Text -gt 27.0 )
	{
		$bemerkung = "Temperature to high!"
	}
	
	elseif ([float]$windowTextBox2.Text -gt 70.0 -or [float]$windowTextBox4.Text -gt 70.0 -or [float]$windowTextBox6.Text -gt 70.0)
	{
		$bemerkung = "Humidty to high!"
	}
	elseif ([float]$windowTextBox.Text -lt 17.0 -or [float]$windowTextBox3.Text -lt 17.0 -or [float]$windowTextBox5.Text -lt 17.0 )
	{
		$bemerkung = "Temperature to low!"
	}
	
	elseif ([float]$windowTextBox2.Text -lt 30.0 -or [float]$windowTextBox4.Text -lt 30 -or [float]$windowTextBox6.Text -lt 30)
	{
		$bemerkung = "Humidty to low!"
	}
	
#----------------------------------------------------------------------------------
# find empty cells

$z= 2400
Do{
    $z; 

    $z = $z + 1

    if ($worksheet.Cells.Item($z,1).Text -match "Sa")
	{
	    $z = $z + 2
	}
    ElseIf($worksheet.Cells.Item($z,1).Text -match "So")
    {
        $z++
    }
} until (($worksheet.Cells.Item($z,1).Text -eq "" ) -or ($worksheet.Cells.Item($z,1).Text -match $Day.Remove(2) -and $worksheet.Cells.Item($z,1).Text -match $DD))

#----------------------------------------------------------------------------------
# Write Excel
    
		$worksheet.Cells.Item($z,1) = $date
	    $worksheet.Cells.Item($z,2) = $zeit
        if ($worksheet.Cells.Item($z,3).Text -eq "")
        {
	        $worksheet.Cells.Item($z,3) = $kuerzel
        }
        else
        {
            $kuerzel = $worksheet.Cells.Item($z,3).Text
        }
	    $worksheet.Cells.Item($z,5) = [float]$windowTextBox.Text
	    $worksheet.Cells.Item($z,6) = [int]$windowTextBox2.Text / 100
	    $worksheet.Cells.Item($z,7) = [float]$windowTextBox3.Text
	    $worksheet.Cells.Item($z,8) = [int]$windowTextBox4.Text / 100
	    $worksheet.Cells.Item($z,10) = "ok"
	    $worksheet.Cells.Item($z,11) = "ok"
	    $worksheet.Cells.Item($z,14) = $bemerkung
	    $worksheet.Cells.Item($z,15) = ""
	    $worksheet.Cells.Item($z,17) = [float]$windowTextBox5.Text
	    $worksheet.Cells.Item($z,18) = [int]$windowTextBox6.Text / 100

sleep 1
clear
#----------------------------------------------------------------------------------
## Speichern + Schliessen
$workbook.Save()
sleep 1
$objexcel.Quit()
clear
   $window.Dispose()
     
  })
 
$window.Controls.Add($windowButton)

[void]$window.ShowDialog()

}
exit


# Ende
#----------------------------------------------------------------------------------
