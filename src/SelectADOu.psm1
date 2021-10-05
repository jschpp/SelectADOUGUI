$UsedProperties = @("Name", "DistinguishedName") 
$OUParams = @{
    Filter      = "*"
    Properties  = $UsedProperties
    SearchScope = "OneLevel"
}

function NewOu {
    [CmdletBinding()]
    param (
        [Parameter()]
        [Object]
        $OU
    )
    $ret = New-Object "System.Windows.Controls.TreeViewItem"
    $ret.Header = $OU.Name
    $ret | Add-Member -TypeName "String" -Value $OU.DistinguishedName -MemberType NoteProperty -Name "DistinguishedName"
    $ret
}

function AddOus {
    param($parent)
    Get-ADOrganizationalUnit @OUParams -SearchBase $parent.DistinguishedName | Select-Object $UsedProperties | ForEach-Object {
        $current = NewOu $_
        AddOus $current
        $parent.AddChild($current)
    }
}


<#
.SYNOPSIS
A graphical selection tool of AD Organisational Units

.DESCRIPTION
This cmdlet will open a WPF Window which can be used to select one OU.

.PARAMETER SearchBase
Searchbase for OU lookup

.EXAMPLE
PS C:\> Select-ADOU -SearchBase "DC=domain,DC=tld"

.NOTES
none.
#>
function Select-ADOU {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidAssignmentToAutomaticVariable', '', Justification = "Correct use of sender in the case of WPF sender")]
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [String]
        $SearchBase = "REMOVED"
    )

    begin {
        [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="600" Height="800" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,0,0,0" Title="OU Auswählen">
	<Grid>
		<TreeView Name="OUView" HorizontalAlignment="Left" BorderBrush="Black" BorderThickness="1" Height="680" VerticalAlignment="Top" Width="580" Margin="10,10,10,10" />
        <Button Content="OK" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="10,700,10,10" Name="OK"/>
	</Grid>
</Window>
"@
        $xamlReader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($xamlReader)

        $treeView = $window.findName('OUView')
        $OKbtn = $window.findName('OK')
        $OKbtn.add_Click( {
                param([System.Object]$sender, [System.Windows.RoutedEventArgs]$e)
                if (-not $treeView.SelectedItem) {
                    return
                }
                $Window.DialogResult = $true
                $window.Close()
            })
    }

    process {
        Get-ADOrganizationalUnit @OUParams -SearchBase $SearchBase | Select-Object $UsedProperties | ForEach-Object {
            $current = NewOu $_
            AddOus $current
            $treeView.Items.Add($current) | Out-Null
        }

        $DialogResult = $window.ShowDialog()
        if ($DialogResult -eq $true) {
            $treeView.SelectedItem.DistinguishedName
        }
    }

    end {}
}