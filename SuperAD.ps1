Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Icon BASE64 Conversion
function ConvertFrom-Base64String {
    param ($base64String)
    $bytes = [System.Convert]::FromBase64String($base64String)
    $stream = New-Object System.IO.MemoryStream(,$bytes)
    return [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::FromStream($stream).GetHicon())
}
$base64Icon = ""  # Inserta aquí tu base64 icon si es necesario

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "SuperAD Administration"
$Form.Size = New-Object System.Drawing.Size(1200, 900)
$Form.StartPosition = 'CenterScreen'
$Form.BackColor = [System.Drawing.Color]::LightSteelBlue

try {
    $Form.Icon = ConvertFrom-Base64String $base64Icon
}
catch {
    Write-Host "The base64 icon is not correct, try to correct it!" -ForegroundColor Red
}

# UI Controls
$UserSearchLabel = New-Object System.Windows.Forms.Label
$UserSearchLabel.Text = "Buscar Usuario:"
$UserSearchLabel.Location = New-Object System.Drawing.Point(10, 10)
$UserSearchLabel.Size = New-Object System.Drawing.Size(100, 20)
$Form.Controls.Add($UserSearchLabel)

$UserSearchBox = New-Object System.Windows.Forms.TextBox
$UserSearchBox.Location = New-Object System.Drawing.Point(120, 10)
$UserSearchBox.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($UserSearchBox)

$UserSearchButton = New-Object System.Windows.Forms.Button
$UserSearchButton.Text = "Buscar"
$UserSearchButton.Location = New-Object System.Drawing.Point(330, 10)
$UserSearchButton.Size = New-Object System.Drawing.Size(75, 23)
$UserSearchButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($UserSearchButton)

$UserListBox = New-Object System.Windows.Forms.ListBox
$UserListBox.Location = New-Object System.Drawing.Point(10, 40)
$UserListBox.Size = New-Object System.Drawing.Size(400, 300)
$Form.Controls.Add($UserListBox)

$GroupSearchLabel = New-Object System.Windows.Forms.Label
$GroupSearchLabel.Text = "Buscar Grupo:"
$GroupSearchLabel.Location = New-Object System.Drawing.Point(10, 350)
$GroupSearchLabel.Size = New-Object System.Drawing.Size(100, 20)
$Form.Controls.Add($GroupSearchLabel)

$GroupSearchBox = New-Object System.Windows.Forms.TextBox
$GroupSearchBox.Location = New-Object System.Drawing.Point(120, 350)
$GroupSearchBox.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($GroupSearchBox)

$GroupSearchButton = New-Object System.Windows.Forms.Button
$GroupSearchButton.Text = "Buscar"
$GroupSearchButton.Location = New-Object System.Drawing.Point(330, 350)
$GroupSearchButton.Size = New-Object System.Drawing.Size(75, 23)
$GroupSearchButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($GroupSearchButton)

$GroupListBox = New-Object System.Windows.Forms.ListBox
$GroupListBox.Location = New-Object System.Drawing.Point(10, 380)
$GroupListBox.Size = New-Object System.Drawing.Size(400, 300)
$Form.Controls.Add($GroupListBox)

$ListBlockedUsersButton = New-Object System.Windows.Forms.Button
$ListBlockedUsersButton.Text = "Listar Usuarios Bloqueados"
$ListBlockedUsersButton.Location = New-Object System.Drawing.Point(430, 40)
$ListBlockedUsersButton.Size = New-Object System.Drawing.Size(200, 23)
$ListBlockedUsersButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($ListBlockedUsersButton)

$UnlockAllUsersButton = New-Object System.Windows.Forms.Button
$UnlockAllUsersButton.Text = "Desbloquear Todos"
$UnlockAllUsersButton.Location = New-Object System.Drawing.Point(430, 70)
$UnlockAllUsersButton.Size = New-Object System.Drawing.Size(200, 23)
$UnlockAllUsersButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($UnlockAllUsersButton)

$ResetPasswordsButton = New-Object System.Windows.Forms.Button
$ResetPasswordsButton.Text = "Restablecer Contraseñas"
$ResetPasswordsButton.Location = New-Object System.Drawing.Point(430, 100)
$ResetPasswordsButton.Size = New-Object System.Drawing.Size(200, 23)
$ResetPasswordsButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($ResetPasswordsButton)

# New Buttons for additional features
$AuditLoginsButton = New-Object System.Windows.Forms.Button
$AuditLoginsButton.Text = "Auditar Logins"
$AuditLoginsButton.Location = New-Object System.Drawing.Point(430, 130)
$AuditLoginsButton.Size = New-Object System.Drawing.Size(200, 23)
$AuditLoginsButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($AuditLoginsButton)

$ManagePasswordPoliciesButton = New-Object System.Windows.Forms.Button
$ManagePasswordPoliciesButton.Text = "Gestionar Políticas de Contraseña"
$ManagePasswordPoliciesButton.Location = New-Object System.Drawing.Point(430, 160)
$ManagePasswordPoliciesButton.Size = New-Object System.Drawing.Size(200, 23)
$ManagePasswordPoliciesButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($ManagePasswordPoliciesButton)

$MassCreateAccountsButton = New-Object System.Windows.Forms.Button
$MassCreateAccountsButton.Text = "Crear Cuentas Masivamente"
$MassCreateAccountsButton.Location = New-Object System.Drawing.Point(430, 190)
$MassCreateAccountsButton.Size = New-Object System.Drawing.Size(200, 23)
$MassCreateAccountsButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($MassCreateAccountsButton)

$GenerateTemplatesButton = New-Object System.Windows.Forms.Button
$GenerateTemplatesButton.Text = "Generar Plantillas"
$GenerateTemplatesButton.Location = New-Object System.Drawing.Point(430, 220)
$GenerateTemplatesButton.Size = New-Object System.Drawing.Size(200, 23)
$GenerateTemplatesButton.BackColor = [System.Drawing.Color]::GhostWhite
$Form.Controls.Add($GenerateTemplatesButton)

# Event Handlers
$UserSearchButton.Add_Click({
    $searchQuery = $UserSearchBox.Text
    $UserListBox.Items.Clear()
    
    if ($searchQuery) {
        try {
            $filter = "Name -like '*$searchQuery*'"
            Write-Host "Buscando usuarios con el filtro: $filter"
            $users = Get-ADUser -Filter $filter -Property Name
            if ($users.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No se encontraron usuarios.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            foreach ($user in $users) {
                $UserListBox.Items.Add($user.Name)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al buscar usuarios: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingrese un texto de búsqueda.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$GroupSearchButton.Add_Click({
    $searchQuery = $GroupSearchBox.Text
    $GroupListBox.Items.Clear()
    
    if ($searchQuery) {
        try {
            $filter = "Name -like '*$searchQuery*'"
            Write-Host "Buscando grupos con el filtro: $filter"
            $groups = Get-ADGroup -Filter $filter -Property Name
            if ($groups.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No se encontraron grupos.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            foreach ($group in $groups) {
                $GroupListBox.Items.Add($group.Name)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al buscar grupos: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingrese un texto de búsqueda.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$ListBlockedUsersButton.Add_Click({
    $UserListBox.Items.Clear()
    try {
        $blockedUsers = Search-ADAccount -LockedOut -UsersOnly
        if ($blockedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No se encontraron usuarios bloqueados.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        foreach ($user in $blockedUsers) {
            $UserListBox.Items.Add($user.Name)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al listar usuarios bloqueados: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$UnlockAllUsersButton.Add_Click({
    try {
        $lockedUsers = Search-ADAccount -LockedOut -UsersOnly
        foreach ($user in $lockedUsers) {
            Unlock-ADAccount -Identity $user.SamAccountName
        }
        [System.Windows.Forms.MessageBox]::Show("Todos los usuarios bloqueados han sido desbloqueados.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al desbloquear usuarios: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$ResetPasswordsButton.Add_Click({
    $user = [System.Windows.Forms.InputBox]::Show("Ingrese el nombre de usuario para restablecer la contraseña:", "Restablecer Contraseñas", "")
    if ($user) {
        try {
            $newPassword = [System.Windows.Forms.InputBox]::Show("Ingrese la nueva contraseña:", "Restablecer Contraseñas", "")
            if ($newPassword) {
                Set-ADAccountPassword -Identity $user -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force) -Reset
                [System.Windows.Forms.MessageBox]::Show("Contraseña restablecida con éxito.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("No se ingresó una nueva contraseña.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al restablecer la contraseña: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No se ingresó el nombre de usuario.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$AuditLoginsButton.Add_Click({
# Función para mostrar la ventana de entrada del nombre de usuario
# Función para mostrar la ventana de entrada del nombre de usuario
function Show-UserInputDialog {
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Ingresar Nombre de Usuario"
    $inputForm.Size = New-Object System.Drawing.Size(300, 150)
    $inputForm.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Ingrese el nombre de usuario:"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(260, 20)
    $inputForm.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $inputForm.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Aceptar"
    $okButton.Location = New-Object System.Drawing.Point(100, 80)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.Controls.Add($okButton)

    $inputForm.AcceptButton = $okButton

    if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}

# Función para mostrar información de inicio y cierre de sesión en tiempo real
function Show-RealTimeUserInfo {
    param (
        [string]$Username
    )

    if (-not $Username) {
        [System.Windows.Forms.MessageBox]::Show("El nombre de usuario no puede estar vacío.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    try {
        # Preparar la ventana para mostrar la información
        $infoForm = New-Object System.Windows.Forms.Form
        $infoForm.Text = "Información de Inicio y Cierre de Sesión en Tiempo Real"
        $infoForm.Size = New-Object System.Drawing.Size(800, 600)
        $infoForm.StartPosition = 'CenterScreen'

        $eventTextBox = New-Object System.Windows.Forms.TextBox
        $eventTextBox.Location = New-Object System.Drawing.Point(10, 10)
        $eventTextBox.Size = New-Object System.Drawing.Size(760, 500)
        $eventTextBox.Multiline = $true
        $eventTextBox.ScrollBars = 'Vertical'
        $eventTextBox.ReadOnly = $true
        $infoForm.Controls.Add($eventTextBox)

        # Mostrar la ventana
        $infoForm.Show()

        # Crear un scriptblock para monitorear eventos en segundo plano
        $scriptBlock = {
            param ($username, $textBox, $form)
            $filter = @{'LogName'='Security'; 'Id'=4624,4647}  # IDs para inicio de sesión y cierre de sesión

            while ($form.Visible) {
                try {
                    # Filtrar eventos relevantes
                    $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop
                    foreach ($event in $events) {
                        $message = $event.Message
                        if ($message -like "*$username*") {
                            $eventDate = $event.TimeCreated.ToString('dd/MM/yyyy - HH:mm')
                            if ($event.Id -eq 4624) {
                                $eventType = 'Log in'
                            } elseif ($event.Id -eq 4647) {
                                $eventType = 'Log off'
                            } else {
                                continue
                            }
                            $updateText = "User: $username, $eventType $eventDate`r`n"
                            $textBox.Invoke([action] { $textBox.AppendText($updateText) })
                        }
                    }
                } catch {
                    $textBox.Invoke([action] { $textBox.AppendText("Error al obtener la información: $_`r`n") })
                }

                Start-Sleep -Seconds 10
            }
        }

        # Iniciar el trabajo en segundo plano
        Start-Job -ScriptBlock $scriptBlock -ArgumentList $Username, $eventTextBox, $infoForm

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al iniciar el monitoreo: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Obtener el nombre de usuario mediante la ventana de entrada
$username = Show-UserInputDialog

# Si se ingresó un nombre de usuario, empezar a monitorear eventos en tiempo real
if ($username) {
    Show-RealTimeUserInfo -Username $username
}
})

$ManagePasswordPoliciesButton.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Funcionalidad de gestión de políticas de contraseña aún no implementada.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$MassCreateAccountsButton.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Funcionalidad para crear cuentas masivamente aún no implementada.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$GenerateTemplatesButton.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Funcionalidad para generar plantillas aún no implementada.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$UserSearchButton.Add_Click({
    $searchQuery = $UserSearchBox.Text
    $UserListBox.Items.Clear()
    
    if ($searchQuery) {
        try {
            $filter = "Name -like '*$searchQuery*'"
            Write-Host "Buscando usuarios con el filtro: $filter"
            $users = Get-ADUser -Filter $filter -Property Name
            if ($users.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No se encontraron usuarios.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            foreach ($user in $users) {
                $UserListBox.Items.Add($user.Name)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al buscar usuarios: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingrese un texto de búsqueda.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$GroupSearchButton.Add_Click({
    $searchQuery = $GroupSearchBox.Text
    $GroupListBox.Items.Clear()
    
    if ($searchQuery) {
        try {
            $filter = "Name -like '*$searchQuery*'"
            Write-Host "Buscando grupos con el filtro: $filter"
            $groups = Get-ADGroup -Filter $filter -Property Name
            if ($groups.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No se encontraron grupos.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            foreach ($group in $groups) {
                $GroupListBox.Items.Add($group.Name)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al buscar grupos: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingrese un texto de búsqueda.", "Advertencia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

$ListBlockedUsersButton.Add_Click({
    $UserListBox.Items.Clear()
    try {
        $blockedUsers = Search-ADAccount -LockedOut -UsersOnly
        if ($blockedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No se encontraron usuarios bloqueados.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        foreach ($user in $blockedUsers) {
            $UserListBox.Items.Add($user.Name)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al buscar usuarios bloqueados: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$UnlockAllUsersButton.Add_Click({
    try {
        $lockedUsers = Search-ADAccount -LockedOut -UsersOnly
        foreach ($user in $lockedUsers) {
            Unlock-ADAccount -Identity $user.SamAccountName
        }
        [System.Windows.Forms.MessageBox]::Show("Todas las cuentas bloqueadas han sido desbloqueadas.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al desbloquear cuentas: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$ResetPasswordsButton.Add_Click({
    $newPassword = [System.Windows.Forms.InputBox]::Show("Ingrese la nueva contraseña para todos los usuarios:", "Restablecer Contraseñas")
    if ($newPassword) {
        try {
            $users = Get-ADUser -Filter * -Property SamAccountName
            foreach ($user in $users) {
                Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword (ConvertTo-SecureString $newPassword -AsPlainText -Force) -Reset
            }
            [System.Windows.Forms.MessageBox]::Show("Contraseñas restablecidas exitosamente para todos los usuarios.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al restablecer contraseñas: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$UserListBox.Add_DoubleClick({
    if ($UserListBox.SelectedItem) {
        $selectedUserName = $UserListBox.SelectedItem
        try {
            $selectedUser = Get-ADUser -Filter { Name -eq $selectedUserName } -Property *
            ShowUserProperties($selectedUser)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al obtener propiedades del usuario: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$GroupListBox.Add_DoubleClick({
    if ($GroupListBox.SelectedItem) {
        $selectedGroupName = $GroupListBox.SelectedItem
        try {
            $selectedGroup = Get-ADGroup -Filter { Name -eq $selectedGroupName } -Property *
            ShowGroupProperties($selectedGroup)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al obtener propiedades del grupo: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

function Show-InputForm {
    param (
        [string]$Prompt,
        [string]$Title
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(260, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Aceptar"
    $button.Location = New-Object System.Drawing.Point(100, 80)
    $button.Size = New-Object System.Drawing.Size(75, 23)
    $button.Add_Click({
        $form.Tag = $textBox.Text
        $form.Close()
    })
    $form.Controls.Add($button)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Show-PasswordForm {
    param (
        [string]$Prompt,
        [string]$Title
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(260, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.UseSystemPasswordChar = $true  # Ocultar el texto ingresado
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Aceptar"
    $button.Location = New-Object System.Drawing.Point(100, 80)
    $button.Size = New-Object System.Drawing.Size(75, 23)
    $button.Add_Click({
        $form.Tag = $textBox.Text
        $form.Close()
    })
    $form.Controls.Add($button)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function ShowUserProperties {
    param ($user)

    $userForm = New-Object System.Windows.Forms.Form
    $userForm.Text = "Propiedades del Usuario"
    $userForm.Size = New-Object System.Drawing.Size(500, 400)
    $userForm.StartPosition = 'CenterScreen'

    $labels = @(
        "Nombre: $($user.Name)",
        "Nombre de Usuario: $($user.SamAccountName)",
        "Correo Electrónico: $($user.EmailAddress)",
        "Teléfono: $($user.TelephoneNumber)",
        "Departamento: $($user.Department)",
        "Fecha de Creación: $($user.WhenCreated)",
        "Último Inicio de Sesión: $($user.LastLogonDate)"
    )
    
    $lockedOut = $user.LockedOut
    $enabled = $user.Enabled
    $statusLabels = @(
        "Estado de la Cuenta:" 
        if ($enabled) {"Habilitada"} else {"Deshabilitada"},
        "Estado de Bloqueo:"
        if ($lockedOut) {"Bloqueada"} else {"No Bloqueada"}
    )
    
    $yPos = 10
    foreach ($labelText in $labels) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelText
        $label.Location = New-Object System.Drawing.Point(10, $yPos)
        $label.Size = New-Object System.Drawing.Size(460, 20)
        $userForm.Controls.Add($label)
        $yPos += 30
    }

    $statusYPos = $yPos
    for ($i = 0; $i -lt $statusLabels.Length; $i += 2) {
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = $statusLabels[$i]
        $titleLabel.Location = New-Object System.Drawing.Point(10, $statusYPos)
        $titleLabel.Size = New-Object System.Drawing.Size(200, 20)
        $userForm.Controls.Add($titleLabel)

        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Text = $statusLabels[$i + 1]
        $statusLabel.Location = New-Object System.Drawing.Point(220, $statusYPos)
        $statusLabel.Size = New-Object System.Drawing.Size(200, 20)
        
        if ($statusLabels[$i + 1] -eq "Habilitada" -or $statusLabels[$i + 1] -eq "No Bloqueada") {
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
        } elseif ($statusLabels[$i + 1] -eq "Deshabilitada" -or $statusLabels[$i + 1] -eq "Bloqueada") {
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        } else {
            $statusLabel.ForeColor = [System.Drawing.Color]::Black
        }

        $userForm.Controls.Add($statusLabel)
        $statusYPos += 30
    }

    $toggleButton = New-Object System.Windows.Forms.Button
    $toggleButton.Text = if ($enabled) { "Deshabilitar" } else { "Habilitar" }
    $toggleButton.Location = New-Object System.Drawing.Point(10, $statusYPos)
    $toggleButton.Size = New-Object System.Drawing.Size(100, 23)
    $toggleButton.Add_Click({
        try {
            if ($enabled) {
                Disable-ADAccount -Identity $user.SamAccountName
                [System.Windows.Forms.MessageBox]::Show("Cuenta deshabilitada exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                Enable-ADAccount -Identity $user.SamAccountName
                [System.Windows.Forms.MessageBox]::Show("Cuenta habilitada exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            ShowUserProperties (Get-ADUser -Identity $user.SamAccountName -Property *)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al cambiar el estado de la cuenta: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $userForm.Controls.Add($toggleButton)

    $unlockButton = New-Object System.Windows.Forms.Button
    $unlockButton.Text = "Desbloquear"
    $unlockButton.Location = New-Object System.Drawing.Point(120, $statusYPos)
    $unlockButton.Size = New-Object System.Drawing.Size(100, 23)
    $unlockButton.Add_Click({
        try {
            Unlock-ADAccount -Identity $user.SamAccountName
            [System.Windows.Forms.MessageBox]::Show("Cuenta desbloqueada exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            ShowUserProperties (Get-ADUser -Identity $user.SamAccountName -Property *)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al desbloquear cuenta: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $userForm.Controls.Add($unlockButton)

    $resetPasswordButton = New-Object System.Windows.Forms.Button
    $resetPasswordButton.Text = "Restablecer Contraseña"
    $resetPasswordButton.Location = New-Object System.Drawing.Point(230, $statusYPos)
    $resetPasswordButton.Size = New-Object System.Drawing.Size(150, 23)
    $resetPasswordButton.Add_Click({
        $newPassword = Show-PasswordForm -Prompt "Ingrese la nueva contraseña:" -Title "Restablecer Contraseña"
        if ($newPassword) {
            try {
                Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword (ConvertTo-SecureString $newPassword -AsPlainText -Force)
                [System.Windows.Forms.MessageBox]::Show("Contraseña restablecida exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error al restablecer contraseña: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $userForm.Controls.Add($resetPasswordButton)
    $userForm.Add_Shown({$userForm.Activate()})
    [void]$userForm.ShowDialog()
}


function ShowGroupProperties {
    param ($group)

    if (-not $group -or $group.GetType().Name -ne 'ADGroup') {
        [System.Windows.Forms.MessageBox]::Show("El objeto del grupo no es válido.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $groupForm = New-Object System.Windows.Forms.Form
    $groupForm.Text = "Propiedades del Grupo"
    $groupForm.Size = New-Object System.Drawing.Size(600, 400)
    $groupForm.StartPosition = 'CenterScreen'

    $labels = @(
        "Nombre: $($group.Name)",
        "Nombre de Grupo: $($group.SamAccountName)",
        "Descripción: $($group.Description)",
        "Fecha de Creación: $($group.WhenCreated)"
    )

    $yPos = 10
    foreach ($labelText in $labels) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelText
        $label.Location = New-Object System.Drawing.Point(10, $yPos)
        $label.Size = New-Object System.Drawing.Size(560, 20)
        $groupForm.Controls.Add($label)
        $yPos += 30
    }
    $modifyButton = New-Object System.Windows.Forms.Button
    $modifyButton.Text = "Modificar Grupo"
    $modifyButton.Location = New-Object System.Drawing.Point(10, $yPos)
    $modifyButton.Size = New-Object System.Drawing.Size(150, 23)
    $modifyButton.Add_Click({
        ShowGroupModificationForm -group $group
    })
    $groupForm.Controls.Add($modifyButton)

    $groupForm.ShowDialog()
}

function ShowGroupModificationForm {
    param ($group)

    # Validar que el objeto $group sea válido
    if (-not $group -or $group.GetType().Name -ne 'ADGroup') {
        [System.Windows.Forms.MessageBox]::Show("El objeto del grupo no es válido.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Crear el formulario de modificación del grupo
    $modifyForm = New-Object System.Windows.Forms.Form
    $modifyForm.Text = "Modificar Grupo"
    $modifyForm.Size = New-Object System.Drawing.Size(750, 500)
    $modifyForm.StartPosition = 'CenterScreen'

    # Crear el CheckedListBox de usuarios actuales
    $userCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $userCheckedListBox.Location = New-Object System.Drawing.Point(10, 10)
    $userCheckedListBox.Size = New-Object System.Drawing.Size(300, 200)
    $modifyForm.Controls.Add($userCheckedListBox)

    # Crear el CheckedListBox de usuarios disponibles
    $availableUsersCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $availableUsersCheckedListBox.Location = New-Object System.Drawing.Point(320, 10)
    $availableUsersCheckedListBox.Size = New-Object System.Drawing.Size(300, 200)
    $modifyForm.Controls.Add($availableUsersCheckedListBox)

    # Cargar los miembros actuales del grupo en el CheckedListBox
    try {
        $groupMembers = Get-ADGroupMember -Identity $group -ErrorAction Stop
        foreach ($member in $groupMembers) {
            $userCheckedListBox.Items.Add($member.SamAccountName, $true)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al cargar miembros del grupo: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # Cargar usuarios disponibles en el CheckedListBox
    try {
        $allUsers = Get-ADUser -Filter * -Property SamAccountName -ErrorAction Stop
        $groupUserNames = $groupMembers | ForEach-Object { $_.SamAccountName }
        foreach ($user in $allUsers) {
            if (-not ($groupUserNames -contains $user.SamAccountName)) {
                $availableUsersCheckedListBox.Items.Add($user.SamAccountName, $false)
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al cargar usuarios disponibles: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # Crear los controles para añadir usuarios
    $addUserButton = New-Object System.Windows.Forms.Button
    $addUserButton.Text = "Añadir Seleccionados"
    $addUserButton.Location = New-Object System.Drawing.Point(320, 220)
    $addUserButton.Size = New-Object System.Drawing.Size(150, 23)
    $addUserButton.Add_Click({
        $selectedUsers = @($availableUsersCheckedListBox.CheckedItems)
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Seleccione uno o más usuarios para añadir.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } else {
            $usersToAdd = $selectedUsers.Clone()  # Crear una copia de la selección
            foreach ($selectedUser in $usersToAdd) {
                try {
                    $user = Get-ADUser -Identity $selectedUser -ErrorAction Stop
                    Add-ADGroupMember -Identity $group -Members $user -ErrorAction Stop
                    $userCheckedListBox.Items.Add($selectedUser, $true)
                    $availableUsersCheckedListBox.Items.Remove($selectedUser)
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error al añadir usuario '$selectedUser': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Usuarios añadidos correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $modifyForm.Controls.Add($addUserButton)

    # Crear los controles para eliminar usuarios
    $removeUserButton = New-Object System.Windows.Forms.Button
    $removeUserButton.Text = "Eliminar Seleccionados"
    $removeUserButton.Location = New-Object System.Drawing.Point(10, 220)
    $removeUserButton.Size = New-Object System.Drawing.Size(150, 23)
    $removeUserButton.Add_Click({
        $selectedUsers = @($userCheckedListBox.CheckedItems)
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Seleccione uno o más usuarios para eliminar.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } else {
            $usersToRemove = $selectedUsers.Clone()  # Crear una copia de la selección
            foreach ($selectedUser in $usersToRemove) {
                try {
                    $user = Get-ADUser -Identity $selectedUser -ErrorAction Stop
                    Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false -ErrorAction Stop
                    $userCheckedListBox.Items.Remove($selectedUser)
                    $availableUsersCheckedListBox.Items.Add($selectedUser, $false)
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error al eliminar usuario '$selectedUser': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Usuarios eliminados correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $modifyForm.Controls.Add($removeUserButton)

    # Mostrar el formulario de modificación
    $modifyForm.ShowDialog()
}

function ShowGroupProperties {
    param ($group)

    # Validar que el objeto $group sea válido
    if (-not $group -or $group.GetType().Name -ne 'ADGroup') {
        [System.Windows.Forms.MessageBox]::Show("El objeto del grupo no es válido.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Crear el formulario principal
    $groupForm = New-Object System.Windows.Forms.Form
    $groupForm.Text = "Propiedades del Grupo"
    $groupForm.Size = New-Object System.Drawing.Size(600, 400)
    $groupForm.StartPosition = 'CenterScreen'

    # Información del grupo
    $labels = @(
        "Nombre: $($group.Name)",
        "Nombre de Grupo: $($group.SamAccountName)",
        "Descripción: $($group.Description)",
        "Fecha de Creación: $($group.WhenCreated)"
    )

    $yPos = 10
    foreach ($labelText in $labels) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelText
        $label.Location = New-Object System.Drawing.Point(10, $yPos)
        $label.Size = New-Object System.Drawing.Size(560, 20)
        $groupForm.Controls.Add($label)
        $yPos += 30
    }

    # Crear un botón para modificar el grupo
    $modifyButton = New-Object System.Windows.Forms.Button
    $modifyButton.Text = "Modificar Grupo"
    $modifyButton.Location = New-Object System.Drawing.Point(10, $yPos)
    $modifyButton.Size = New-Object System.Drawing.Size(150, 23)
    $modifyButton.Add_Click({
        ShowGroupModificationForm -group $group
    })
    $groupForm.Controls.Add($modifyButton)

    # Mostrar el formulario
    $groupForm.ShowDialog()
}

# Mostrar el formulario principal
[void]$Form.ShowDialog()
