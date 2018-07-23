# Griffin Saiia, James Lipitow
# Career Fair Recruitment Improvement Project
# Data Entry Example

# function to check file existence
function CheckFile($filePath){
        if( Test-Path -Path $filePath ){
            return $TRUE
        }
        else {
			return $FALSE
        }
}

# function that returns the next available row
function Find-Last($objWorksheet){
	$i = 1
	if( [string]::IsNullOrEmpty($objWorksheet.Cells.Item($i,1).text) ){
		Do{
			$i++
		}
        Until( [string]::IsNullOrEmpty($objWorksheet.Cells.Item($i,1).text))
	}
	return $i
}

# function I found online that ensures good release of documents
function Release-Ref ($ref) {
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject( [System.__ComObject]$ref) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
}

# function that encapsulates the data entry
function Log-Data($data){
    $objExcel = new-object -ComObject Excel.Application
	$filePath = "C:\Users\Thinkbox admin\Desktop\recruitement.xlsx"
	$boolean = CheckFile($filePath)
	$i = 1
	if($boolean){
		$objWorkbook = $objExcel.Workbooks.Open($filePath)
		$objWorksheet = $objWorkbook.Worksheets.Item(1)
		$i = Find-Last($objWorksheet)
	}
	else{
		$objWorkbook = $objExcel.Workbooks.Add()
		$objWorksheet = $objWorkbook.Worksheets.Item(1)
	}
	$j = 1
	Foreach ($datum in $data){
		$objWorksheet.Cells.Item($i, $j) = $datum
		$j++
	}
    $j = 1
    $i = 1
    if($boolean){
        $objWorkbook.Save()
    }
    else{
	   $objWorkbook.SaveAs($filePath)
    }
	$objExcel.Workbooks.Close()
    $objExcel.Quit()
}

# ************ main method of sorts ************

# load the Winforms
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
Add-Type -AssemblyName System.Windows.Forms
# create the form
$form = New-Object Windows.Forms.Form
# set the dialog title
$form.text = "Career Fair Recruiting"
# create the label control and set text, size and location
$label1 = New-Object Windows.Forms.Label
$label1.Location = New-Object Drawing.Point 50,30
$label1.Size = New-Object Drawing.Point 200,15
$label1.text = "First Name"
# create TextBox and set text, size and location
$textfield1 = New-Object Windows.Forms.TextBox
$textfield1.Location = New-Object Drawing.Point 50,60
$textfield1.Size = New-Object Drawing.Point 200,30
# create the label control and set text, size and location
$label2 = New-Object Windows.Forms.Label
$label2.Location = New-Object Drawing.Point 50,90
$label2.Size = New-Object Drawing.Point 200,15
$label2.text = "Last name"
# create TextBox and set text, size and location
$textfield2 = New-Object Windows.Forms.TextBox
$textfield2.Location = New-Object Drawing.Point 50,120
$textfield2.Size = New-Object Drawing.Point 200,30
# create the label control and set text, size and location
$label3 = New-Object Windows.Forms.Label
$label3.Location = New-Object Drawing.Point 50,150
$label3.Size = New-Object Drawing.Point 200,15
$label3.text = "Email"
# create TextBox and set text, size and location
$textfield3 = New-Object Windows.Forms.TextBox
$textfield3.Location = New-Object Drawing.Point 50,180
$textfield3.Size = New-Object Drawing.Point 200,30
# create the label control and set text, size and location
$label4 = New-Object Windows.Forms.Label
$label4.Location = New-Object Drawing.Point 50,210
$label4.Size = New-Object Drawing.Point 200,15
$label4.text = "Phone number"
# create TextBox and set text, size and location
$textfield4 = New-Object Windows.Forms.TextBox
$textfield4.Location = New-Object Drawing.Point 50,240
$textfield4.Size = New-Object Drawing.Point 200,30
# create the label control and set text, size and location
$label5 = New-Object Windows.Forms.Label
$label5.Location = New-Object Drawing.Point 50,270
$label5.Size = New-Object Drawing.Point 200,15
$label5.text = "College"
# create TextBox and set text, size and location
$textfield5 = New-Object Windows.Forms.TextBox
$textfield5.Location = New-Object Drawing.Point 50,300
$textfield5.Size = New-Object Drawing.Point 200,30
# create Button and set text and location
$button = New-Object Windows.Forms.Button
$button.text = "Enter"
$button.Location = New-Object Drawing.Point 100,330
# add the controls to the Form
$form.controls.add($label1)
$form.controls.add($textfield1)
$form.controls.add($label2)
$form.controls.add($textfield2)
$form.controls.add($label3)
$form.controls.add($textfield3)
$form.controls.add($label4)
$form.controls.add($textfield4)
$form.controls.add($label5)
$form.controls.add($textfield5)
$form.controls.add($button)
$button.add_Click({
	# creates array to store data
    $data = $textField1.Text, $textField2.Text, $textField3.Text, $textField4.Text, $textField5.Text
    # clears out the form
    $textfield1.Text = ''
    $textfield2.Text = ''
    $textfield3.Text = ''
    $textfield4.Text = ''
    $textfield5.Text = ''
	# logs data
	Log-Data($data)
})
# format window size
$form.Width = 350
$form.Height = 450
# Display the dialog
$form.ShowDialog()
