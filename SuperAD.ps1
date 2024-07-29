Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Icon BASE64 Convertion
function ConvertFrom-Base64String {
    param ($base64String)
    $bytes = [System.Convert]::FromBase64String($base64String)
    $stream = New-Object System.IO.MemoryStream(,$bytes)
    return [System.Drawing.Icon]::FromHandle([System.Drawing.Bitmap]::FromStream($stream).GetHicon())
}
$base64Icon = ""

# Crear el formulario principal
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "SuperAD Administration"
$Form.Size = New-Object System.Drawing.Size(1200, 900)
$Form.StartPosition = 'CenterScreen'
$Form.BackColor = [System.Drawing.Color]::LightSteelBlue

try {
    $Form.Icon = ConvertFrom-Base64String $base64Icon
}
catch {
    Write-Host "The base64 icon is not correct, try to correct it!" -ForegroundColor Re
}
# Crear controles para buscar usuarios
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

# Crear controles para buscar grupos
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

# Botones adicionales para acciones masivas
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

# Evento de clic en botón de búsqueda de usuarios
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

# Evento de clic en botón de búsqueda de grupos
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

# Evento de clic en el botón de listar usuarios bloqueados
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

# Evento de clic en el botón de desbloquear todos los usuarios
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

# Evento de clic en el botón de restablecer contraseñas
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

# Evento de doble clic en un usuario de la lista
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

# Evento de doble clic en un grupo de la lista
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

# Función para mostrar un formulario de entrada
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

# Función para mostrar un formulario de entrada con texto oculto (contraseña)
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

# Función para mostrar propiedades del usuario
function ShowUserProperties {
    param ($user)

    $userForm = New-Object System.Windows.Forms.Form
    $userForm.Text = "Propiedades del Usuario"
    $userForm.Size = New-Object System.Drawing.Size(500, 400)
    $userForm.StartPosition = 'CenterScreen'

    # Información del usuario
    $labels = @(
        "Nombre: $($user.Name)",
        "Nombre de Usuario: $($user.SamAccountName)",
        "Correo Electrónico: $($user.EmailAddress)",
        "Teléfono: $($user.TelephoneNumber)",
        "Departamento: $($user.Department)",
        "Fecha de Creación: $($user.WhenCreated)",
        "Último Inicio de Sesión: $($user.LastLogonDate)"
    )
    
    # Estado de la cuenta y bloqueo
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

    # Estado de la cuenta y bloqueo
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

    # Botón para habilitar o deshabilitar
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
            # Actualizar la información del usuario
            ShowUserProperties (Get-ADUser -Identity $user.SamAccountName -Property *)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al cambiar el estado de la cuenta: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $userForm.Controls.Add($toggleButton)

    # Botón para desbloquear
    $unlockButton = New-Object System.Windows.Forms.Button
    $unlockButton.Text = "Desbloquear"
    $unlockButton.Location = New-Object System.Drawing.Point(120, $statusYPos)
    $unlockButton.Size = New-Object System.Drawing.Size(100, 23)
    $unlockButton.Add_Click({
        try {
            Unlock-ADAccount -Identity $user.SamAccountName
            [System.Windows.Forms.MessageBox]::Show("Cuenta desbloqueada exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            # Actualizar la información del usuario
            ShowUserProperties (Get-ADUser -Identity $user.SamAccountName -Property *)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al desbloquear cuenta: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $userForm.Controls.Add($unlockButton)

    # Botón para restablecer contraseña
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

    # Botón para crear nuevo usuario
    $createUserButtonYPos = $statusYPos + 40
    $createUserButton = New-Object System.Windows.Forms.Button
    $createUserButton.Text = "Crear Nuevo Usuario"
    $createUserButton.Location = New-Object System.Drawing.Point(10, $createUserButtonYPos)
    $createUserButton.Size = New-Object System.Drawing.Size(150, 23)
    $createUserButton.Add_Click({
        $newUserName = Show-InputForm -Prompt "Ingrese el nombre del nuevo usuario:" -Title "Crear Nuevo Usuario"
        
        if ($newUserName) {
            $newUserPassword = Show-PasswordForm -Prompt "Ingrese la contraseña para el nuevo usuario:" -Title "Crear Nuevo Usuario"

            if ($newUserPassword) {
                try {
                    # Extraer el dominio del UserPrincipalName del usuario actual
                    $domain = $user.UserPrincipalName.Split('@')[1]

                    # Definir los atributos para el nuevo usuario basados en el usuario actual, excluyendo el nombre y la contraseña
                    $newUserParams = @{
                        Name = $newUserName
                        GivenName = $user.GivenName
                        Surname = $user.Surname
                        SamAccountName = $newUserName
                        UserPrincipalName = "$newUserName@$domain"  # Usar el dominio del usuario actual
                        Path = $user.DistinguishedName  # Usar la misma unidad organizativa que el usuario actual
                        AccountPassword = (ConvertTo-SecureString $newUserPassword -AsPlainText -Force)  # Usar la contraseña proporcionada
                        Enabled = $true
                        EmailAddress = $user.EmailAddress
                        Department = $user.Department
                        Title = $user.Title
                        Company = $user.Company
                        Office = $user.Office
                        Manager = $user.Manager
                        StreetAddress = $user.StreetAddress
                        City = $user.City
                        State = $user.State
                        PostalCode = $user.PostalCode
                        Country = $user.Country
                        Description = $user.Description
                    }

                    # Crear el nuevo usuario
                    New-ADUser @newUserParams

                    [System.Windows.Forms.MessageBox]::Show("Nuevo usuario creado exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error al crear nuevo usuario: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        }
    })
    $userForm.Controls.Add($createUserButton)

    # Mostrar el formulario
    $userForm.Add_Shown({$userForm.Activate()})
    [void]$userForm.ShowDialog()
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
