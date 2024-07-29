Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "SuperAD Administration"
$Form.Size = New-Object System.Drawing.Size(1200,900)
$Form.StartPosition = 'CenterScreen'


$Form.ShowDialog()