
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(683,498)
$Form.text                       = "VM Cloner - Daniel Beilin"
$Form.TopMost                    = $false

$VCenterLbl                      = New-Object system.Windows.Forms.Label
$VCenterLbl.text                 = "VCenter"
$VCenterLbl.AutoSize             = $true
$VCenterLbl.width                = 25
$VCenterLbl.height               = 10
$VCenterLbl.location             = New-Object System.Drawing.Point(26,83)
$VCenterLbl.Font                 = New-Object System.Drawing.Font('Arial',12)

$OSLbl                           = New-Object system.Windows.Forms.Label
$OSLbl.text                      = "OS"
$OSLbl.AutoSize                  = $true
$OSLbl.width                     = 25
$OSLbl.height                    = 10
$OSLbl.location                  = New-Object System.Drawing.Point(26,106)
$OSLbl.Font                      = New-Object System.Drawing.Font('Arial',12)

$TemplateLbl                     = New-Object system.Windows.Forms.Label
$TemplateLbl.text                = "Template"
$TemplateLbl.AutoSize            = $true
$TemplateLbl.width               = 25
$TemplateLbl.height              = 10
$TemplateLbl.location            = New-Object System.Drawing.Point(26,129)
$TemplateLbl.Font                = New-Object System.Drawing.Font('Arial',12)

$csvLbl                          = New-Object system.Windows.Forms.Label
$csvLbl.text                     = "CSV"
$csvLbl.AutoSize                 = $true
$csvLbl.width                    = 25
$csvLbl.height                   = 10
$csvLbl.location                 = New-Object System.Drawing.Point(26,198)
$csvLbl.Font                     = New-Object System.Drawing.Font('Arial',12)

$DatastoreLbl                    = New-Object system.Windows.Forms.Label
$DatastoreLbl.text               = "Datastore"
$DatastoreLbl.AutoSize           = $true
$DatastoreLbl.width              = 25
$DatastoreLbl.height             = 10
$DatastoreLbl.location           = New-Object System.Drawing.Point(26,175)
$DatastoreLbl.Font               = New-Object System.Drawing.Font('Arial',12)

$CPULbl                          = New-Object system.Windows.Forms.Label
$CPULbl.text                     = "CPU"
$CPULbl.AutoSize                 = $true
$CPULbl.width                    = 25
$CPULbl.height                   = 10
$CPULbl.location                 = New-Object System.Drawing.Point(26,220)
$CPULbl.Font                     = New-Object System.Drawing.Font('Arial',12)

$RAMLbl                          = New-Object system.Windows.Forms.Label
$RAMLbl.text                     = "RAM"
$RAMLbl.AutoSize                 = $true
$RAMLbl.width                    = 25
$RAMLbl.height                   = 10
$RAMLbl.location                 = New-Object System.Drawing.Point(26,243)
$RAMLbl.Font                     = New-Object System.Drawing.Font('Arial',12)

$NotesLbl                        = New-Object system.Windows.Forms.Label
$NotesLbl.text                   = "Notes"
$NotesLbl.AutoSize               = $true
$NotesLbl.width                  = 25
$NotesLbl.height                 = 10
$NotesLbl.location               = New-Object System.Drawing.Point(26,312)
$NotesLbl.Font                   = New-Object System.Drawing.Font('Arial',12)

$VCenterTB                       = New-Object system.Windows.Forms.TextBox
$VCenterTB.multiline             = $false
$VCenterTB.width                 = 338
$VCenterTB.height                = 20
$VCenterTB.location              = New-Object System.Drawing.Point(195,76)
$VCenterTB.Font                  = New-Object System.Drawing.Font('Arial',12)

$csvTB                           = New-Object system.Windows.Forms.TextBox
$csvTB.multiline                 = $false
$csvTB.width                     = 338
$csvTB.height                    = 20
$csvTB.enabled                   = $false
$csvTB.location                  = New-Object System.Drawing.Point(195,196)
$csvTB.Font                      = New-Object System.Drawing.Font('Arial',12)

$CPU_TB                          = New-Object system.Windows.Forms.TextBox
$CPU_TB.multiline                = $false
$CPU_TB.width                    = 338
$CPU_TB.height                   = 20
$CPU_TB.enabled                  = $false
$CPU_TB.location                 = New-Object System.Drawing.Point(195,220)
$CPU_TB.Font                     = New-Object System.Drawing.Font('Arial',12)

$RAM_TB                          = New-Object system.Windows.Forms.TextBox
$RAM_TB.multiline                = $false
$RAM_TB.text                     = "In GBs"
$RAM_TB.width                    = 338
$RAM_TB.height                   = 20
$RAM_TB.enabled                  = $false
$RAM_TB.location                 = New-Object System.Drawing.Point(195,244)
$RAM_TB.Font                     = New-Object System.Drawing.Font('Arial',12)

$NotesTB                         = New-Object system.Windows.Forms.TextBox
$NotesTB.multiline               = $false
$NotesTB.width                   = 338
$NotesTB.height                  = 20
$NotesTB.enabled                 = $false
$NotesTB.location                = New-Object System.Drawing.Point(195,316)
$NotesTB.Font                    = New-Object System.Drawing.Font('Arial',12)

$OS_Menu                         = New-Object system.Windows.Forms.ComboBox
$OS_Menu.width                   = 338
$OS_Menu.height                  = 20
$OS_Menu.visible                 = $true
$OS_Menu.enabled                 = $false
@('Windows','CentOS','Red Hat') | ForEach-Object {[void] $OS_Menu.Items.Add($_)}
$OS_Menu.location                = New-Object System.Drawing.Point(195,100)
$OS_Menu.Font                    = New-Object System.Drawing.Font('Arial',12)

$TemplateMenu                    = New-Object system.Windows.Forms.ComboBox
$TemplateMenu.width              = 338
$TemplateMenu.height             = 20
$TemplateMenu.enabled            = $false
$TemplateMenu.location           = New-Object System.Drawing.Point(195,124)
$TemplateMenu.Font               = New-Object System.Drawing.Font('Arial',12)

$DatastoreMenu                   = New-Object system.Windows.Forms.ComboBox
$DatastoreMenu.width             = 338
$DatastoreMenu.height            = 20
$DatastoreMenu.enabled           = $false
$DatastoreMenu.location          = New-Object System.Drawing.Point(195,172)
$DatastoreMenu.Font              = New-Object System.Drawing.Font('Arial',12)

$OptionsLbl                      = New-Object system.Windows.Forms.Label
$OptionsLbl.text                 = "Options"
$OptionsLbl.AutoSize             = $true
$OptionsLbl.width                = 25
$OptionsLbl.height               = 10
$OptionsLbl.location             = New-Object System.Drawing.Point(26,354)
$OptionsLbl.Font                 = New-Object System.Drawing.Font('Arial',12)

$StartVM_CB                      = New-Object system.Windows.Forms.CheckBox
$StartVM_CB.text                 = "Start VM After Creation"
$StartVM_CB.AutoSize             = $false
$StartVM_CB.width                = 424
$StartVM_CB.height               = 20
$StartVM_CB.enabled              = $false
$StartVM_CB.location             = New-Object System.Drawing.Point(195,354)
$StartVM_CB.Font                 = New-Object System.Drawing.Font('Arial',12)

$OKButton                        = New-Object system.Windows.Forms.Button
$OKButton.text                   = "Start Clone"
$OKButton.width                  = 115
$OKButton.height                 = 30
$OKButton.location               = New-Object System.Drawing.Point(12,437)
$OKButton.Font                   = New-Object System.Drawing.Font('Arial',12)

$CancelButton                    = New-Object system.Windows.Forms.Button
$CancelButton.text               = "Cancel"
$CancelButton.width              = 115
$CancelButton.height             = 30
$CancelButton.location           = New-Object System.Drawing.Point(554,437)
$CancelButton.Font               = New-Object System.Drawing.Font('Arial',12)

$ConnectButton                   = New-Object system.Windows.Forms.Button
$ConnectButton.text              = "Connect"
$ConnectButton.width             = 87
$ConnectButton.height            = 21
$ConnectButton.location          = New-Object System.Drawing.Point(535,75)
$ConnectButton.Font              = New-Object System.Drawing.Font('Arial',12)
$ConnectButton.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#7ed321")

$FetchButton                     = New-Object system.Windows.Forms.Button
$FetchButton.text                = "Fetch"
$FetchButton.width               = 87
$FetchButton.height              = 21
$FetchButton.location            = New-Object System.Drawing.Point(536,100)
$FetchButton.Font                = New-Object System.Drawing.Font('Arial',12)

$StatusLbl                       = New-Object system.Windows.Forms.Label
$StatusLbl.text                  = "Connection Status:"
$StatusLbl.AutoSize              = $true
$StatusLbl.width                 = 25
$StatusLbl.height                = 10
$StatusLbl.location              = New-Object System.Drawing.Point(6,477)
$StatusLbl.Font                  = New-Object System.Drawing.Font('Arial',11)

$ConnectedStatus                 = New-Object system.Windows.Forms.Label
$ConnectedStatus.text            = "Idle"
$ConnectedStatus.AutoSize        = $true
$ConnectedStatus.width           = 25
$ConnectedStatus.height          = 10
$ConnectedStatus.location        = New-Object System.Drawing.Point(146,478)
$ConnectedStatus.Font            = New-Object System.Drawing.Font('Arial',11)

$ProvisionedLbl                  = New-Object system.Windows.Forms.Label
$ProvisionedLbl.text             = "Provisioned"
$ProvisionedLbl.AutoSize         = $true
$ProvisionedLbl.width            = 25
$ProvisionedLbl.height           = 10
$ProvisionedLbl.location         = New-Object System.Drawing.Point(425,373)
$ProvisionedLbl.Font             = New-Object System.Drawing.Font('Arial',10)
$ProvisionedLbl.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

$FreeLbl                         = New-Object system.Windows.Forms.Label
$FreeLbl.text                    = "Free"
$FreeLbl.AutoSize                = $true
$FreeLbl.width                   = 25
$FreeLbl.height                  = 10
$FreeLbl.location                = New-Object System.Drawing.Point(425,392)
$FreeLbl.Font                    = New-Object System.Drawing.Font('Arial',10)
$FreeLbl.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#7ed321")

$VLANLbl                         = New-Object system.Windows.Forms.Label
$VLANLbl.text                    = "VLAN"
$VLANLbl.AutoSize                = $true
$VLANLbl.width                   = 25
$VLANLbl.height                  = 10
$VLANLbl.location                = New-Object System.Drawing.Point(26,266)
$VLANLbl.Font                    = New-Object System.Drawing.Font('Arial',12)

$VLAN_CB                         = New-Object system.Windows.Forms.ComboBox
$VLAN_CB.width                   = 338
$VLAN_CB.height                  = 20
$VLAN_CB.enabled                 = $false
$VLAN_CB.location                = New-Object System.Drawing.Point(195,268)
$VLAN_CB.Font                    = New-Object System.Drawing.Font('Arial',12)

$FolderLbl                       = New-Object system.Windows.Forms.Label
$FolderLbl.text                  = "Folder"
$FolderLbl.AutoSize              = $true
$FolderLbl.width                 = 25
$FolderLbl.height                = 10
$FolderLbl.location              = New-Object System.Drawing.Point(26,289)
$FolderLbl.Font                  = New-Object System.Drawing.Font('Arial',12)

$FolderFetchButton               = New-Object system.Windows.Forms.Button
$FolderFetchButton.text          = "Find Folder"
$FolderFetchButton.width         = 87
$FolderFetchButton.height        = 21
$FolderFetchButton.location      = New-Object System.Drawing.Point(536,292)
$FolderFetchButton.Font          = New-Object System.Drawing.Font('Arial',10)

$FolderCB                        = New-Object system.Windows.Forms.ComboBox
$FolderCB.width                  = 338
$FolderCB.height                 = 20
$FolderCB.enabled                = $false
$FolderCB.location               = New-Object System.Drawing.Point(195,292)
$FolderCB.Font                   = New-Object System.Drawing.Font('Arial',12)

$CustomizationCB                 = New-Object system.Windows.Forms.ComboBox
$CustomizationCB.width           = 338
$CustomizationCB.height          = 20
$CustomizationCB.enabled         = $false
$CustomizationCB.location        = New-Object System.Drawing.Point(195,148)
$CustomizationCB.Font            = New-Object System.Drawing.Font('Arial',12)

$CustomizationLbl                = New-Object system.Windows.Forms.Label
$CustomizationLbl.text           = "Custom Template"
$CustomizationLbl.AutoSize       = $true
$CustomizationLbl.width          = 25
$CustomizationLbl.height         = 10
$CustomizationLbl.location       = New-Object System.Drawing.Point(26,152)
$CustomizationLbl.Font           = New-Object System.Drawing.Font('Arial',12)

$StorageLeftLbl                  = New-Object system.Windows.Forms.Label
$StorageLeftLbl.text             = "Storage Left After Clone"
$StorageLeftLbl.AutoSize         = $true
$StorageLeftLbl.width            = 25
$StorageLeftLbl.height           = 10
$StorageLeftLbl.location         = New-Object System.Drawing.Point(425,411)
$StorageLeftLbl.Font             = New-Object System.Drawing.Font('Arial',10)
$StorageLeftLbl.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#d0021b")

$DSInfoLbl                       = New-Object system.Windows.Forms.Label
$DSInfoLbl.text                  = "Data Store Information"
$DSInfoLbl.AutoSize              = $true
$DSInfoLbl.width                 = 25
$DSInfoLbl.height                = 10
$DSInfoLbl.location              = New-Object System.Drawing.Point(462,352)
$DSInfoLbl.Font                  = New-Object System.Drawing.Font('Arial',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Underline))

$ProvisionedLbl2                 = New-Object system.Windows.Forms.Label
$ProvisionedLbl2.AutoSize        = $true
$ProvisionedLbl2.width           = 25
$ProvisionedLbl2.height          = 10
$ProvisionedLbl2.location        = New-Object System.Drawing.Point(581,373)
$ProvisionedLbl2.Font            = New-Object System.Drawing.Font('Arial',10)
$ProvisionedLbl2.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

$FreeLbl2                        = New-Object system.Windows.Forms.Label
$FreeLbl2.AutoSize               = $true
$FreeLbl2.width                  = 25
$FreeLbl2.height                 = 10
$FreeLbl2.location               = New-Object System.Drawing.Point(581,392)
$FreeLbl2.Font                   = New-Object System.Drawing.Font('Arial',10)
$FreeLbl2.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#7ed321")

$StorageLeftLbl2                 = New-Object system.Windows.Forms.Label
$StorageLeftLbl2.AutoSize        = $true
$StorageLeftLbl2.width           = 25
$StorageLeftLbl2.height          = 10
$StorageLeftLbl2.location        = New-Object System.Drawing.Point(581,411)
$StorageLeftLbl2.Font            = New-Object System.Drawing.Font('Arial',10)
$StorageLeftLbl2.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#d0021b")

$loadCSV                         = New-Object system.Windows.Forms.Button
$loadCSV.text                    = "Load"
$loadCSV.width                   = 87
$loadCSV.height                  = 21
$loadCSV.location                = New-Object System.Drawing.Point(536,197)
$loadCSV.Font                    = New-Object System.Drawing.Font('Arial',12)

$calculateStorageBtn             = New-Object system.Windows.Forms.Button
$calculateStorageBtn.text        = "Calculate"
$calculateStorageBtn.width       = 85
$calculateStorageBtn.height      = 41
$calculateStorageBtn.location    = New-Object System.Drawing.Point(451,431)
$calculateStorageBtn.Font        = New-Object System.Drawing.Font('Arial',10)

$Form.controls.AddRange(@($VCenterLbl,$OSLbl,$TemplateLbl,$csvLbl,$DatastoreLbl,$CPULbl,$RAMLbl,$NotesLbl,$VCenterTB,$csvTB,$CPU_TB,$RAM_TB,$NotesTB,$OS_Menu,$TemplateMenu,$DatastoreMenu,$OptionsLbl,$StartVM_CB,$OKButton,$CancelButton,$ConnectButton,$FetchButton,$StatusLbl,$ConnectedStatus,$ProvisionedLbl,$FreeLbl,$VLANLbl,$VLAN_CB,$FolderLbl,$FolderFetchButton,$FolderCB,$CustomizationCB,$CustomizationLbl,$StorageLeftLbl,$DSInfoLbl,$ProvisionedLbl2,$FreeLbl2,$StorageLeftLbl2,$loadCSV,$calculateStorageBtn))

$ConnectButton.Add_MouseClick({ VCenterConnect })
$FetchButton.Add_MouseClick({ FetchOS })
$OKButton.Add_MouseClick({ StartClone })
$CancelButton.Add_MouseClick({ Cancel })
$DatastoreMenu.Add_SelectedValueChanged({ DatastoreChosen })
$FolderFetchButton.Add_MouseClick({ FolderFetch })
$loadCSV.Add_MouseClick({ loadCSV })
$calculateStorageBtn.Add_MouseClick({ CalculateStorage })

function loadCSV {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'CSV Files (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
    }
    $null = $FileBrowser.ShowDialog()
    $csvTB.Text = $FileBrowser.FileName
 }

 function CalculateStorage {
    $NumClones = [int]((Get-Content $csvTB.Text).Length - 1)
    $TemplateName = $TemplateMenu.Text #Template Name
    $TP = Get-Template -Name $TemplateName
    $TemplateSize = [math]::round($tp.extensiondata.storage.perdatastoreusage.committed/1073741824/1) #Get size in an int32 variable to multiply later. We also rounded it.
    $TotalSize = ($NumClones * $TemplateSize)
    $Global:StorageLeft = $script:DS_FreeSpace - $TotalSize
    $StorageLeftLbl2.Text = ($Global:StorageLeft, 'GB')
}
function FolderFetch { 
    $FolderCB.Items.Clear()
    $FolderCB_Contents = $FolderCB.Text
    $FolderList = Get-Folder -Name "*$FolderCB_Contents*"
    foreach ($folder in $FolderList) {$FolderCB.Items.Add($folder)}
}
function DatastoreChosen { 
    $DS_Name = Get-Datastore $DatastoreMenu.SelectedItem | Get-View | Select-Object -ExpandProperty summary
    $DS_Capacity = [math]::round($DS_Name.Capacity/1GB)
    $script:DS_FreeSpace = [math]::round($DS_Name.FreeSpace/1GB) #Use the $script variable to use $DS_FreeSpace in another function
    $ProvisionedLbl2.Text = $DS_Capacity, 'GB'
    $FreeLbl2.Text = $DS_FreeSpace, 'GB'
    }
function Cancel { Disconnect-VIServer -Server $VCenterTB.Text -ErrorAction SilentlyContinue -Confirm:$false
    $Form.Close() }
    
function FetchOS { 
    #Clear combox each time a user fetches the OS information
    $CustomizationCB.Items.Clear()
    $TemplateMenu.Items.Clear()
    $DatastoreMenu.Items.Clear()

    #Enable fields once OS is chosen
    $TemplateMenu.Enabled = $true
    $CustomizationCB.Enabled = $true
    $DatastoreMenu.Enabled = $true
    $CPU_TB.Enabled = $true
    $RAM_TB.Enabled = $true
    $VLAN_CB.Enabled = $true
    $FolderCB.Enabled = $true
    $NotesTB.Enabled = $true
    $StartVM_CB.Enabled = $true
    $csvTB.Enabled = $true

    #Populates menus with relavent data
    #Windows
    if ($OS_Menu.SelectedItem -eq 'Windows') 
        {$template_list = Get-Template -Name *Win*
        $datastore_list = Get-Datastore -Name *Win*
        $CustomTemplates = Get-OSCustomizationSpec | Where-Object {$_.Name -like "*WIN*"}}

    #Centos
    elseif ($OS_Menu.SelectedItem -eq 'CentOS')
        {$template_list = Get-Template -Name *CentOS*
        $datastore_list = Get-Datastore -Name *CentOS*
        $CustomTemplates = Get-OSCustomizationSpec | Where-Object {$_.Name -like "*CentOS*"}}

    #Redhat
    elseif ($OS_Menu.SelectedItem -eq 'Red Hat')
        {$template_list = Get-Template -Name *RHEL*
        $datastore_list = Get-Datastore -Name *CentOS*
        $CustomTemplates = Get-OSCustomizationSpec | Where-Object {$_.Name -like "*RH*"}}

    foreach ($temp in $template_list) {$TemplateMenu.Items.Add($temp)}
    foreach ($ds in $datastore_list) {$DatastoreMenu.Items.Add($ds)}
    foreach ($template in $CustomTemplates) {$CustomizationCB.Items.Add($template)}
}

function VCenterConnect { 
    if ($VCenterTB.Text -ne '') {
        try {
            Connect-VIServer $VCenterTB.Text -ErrorAction Stop
            $ConnectedStatus.ForeColor = "#7ed321"
            $ConnectedStatus.Text = "Connected to " + $VCenterTB.Text
            $OS_Menu.Enabled = $true
            VLANs_List
        }
        catch {
            $ConnectedStatus.ForeColor = "#d0021b"
            $ConnectedStatus.Text = "Failed to connect"
        }
    }
}

function VLANs_List {
    $VLANs = Get-VirtualPortGroup -Datacenter "Holon Datacenter" | Where-Object {$_.Name -like "*VLAN*"} | Select-Object -unique
    foreach ($VLAN in $VLANs) {$VLAN_CB.Items.Add($VLAN)}
    }

function StartClone {
    $VM_List = Import-Csv $csvTB.Text
    $numClones = [int]((Get-Content $csvTB.Text).Length)
    $vmh = Get-VMHost
    $NewParameters = @{
        # Name                = ''
        Template            = $TemplateMenu.Text
        Datastore           = $DatastoreMenu.Text
        DiskStorageFormat   = 'Thin'
        Location            = $FolderCB.Text
        OSCustomizationSpec = $CustomizationCB.Text
        VMHost              = Get-Random -InputObject $vmh
        Server              = $VCenterTB.Text
        RunAsync            = $true
    }

    $SetParameters = @{
        NumCpu              = $CPU_TB.Text
        MemoryGB            = $RAM_TB.Text
        Notes               = $NotesTB.Text
        Confirm             = $false
    }

    $taskList = if ($NumClones -gt 0) {
        $VM_List | ForEach-Object {
            $NewParameters['Name'] = "$($_.Hostname)"
            Get-OSCustomizationSpec -name $CustomizationCB.Text | Get-OSCustomizationNICMapping | Set-OSCustomizationNICMapping -IPMode UseStaticIP -IPAddress "$($_.IP)" -SubNetMask "$($_.Subnet)" -DefaultGateway "$($_.Gateway)" -Dns "DNS1","DNS2"
            New-VM @NewParameters
        }
    }
    $newVM = $taskList | Wait-Task -ErrorAction SilentlyContinue
    $newVM | Set-VM @SetParameters
    $newVM | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName $VLAN_CB.Text -Confirm:$false
    if ($StartVM_CB.Checked -eq $true) {$newVM | Start-VM }
 }


[void]$Form.ShowDialog()
