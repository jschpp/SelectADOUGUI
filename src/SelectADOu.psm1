function SplitDn {
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$dn,

        [string]$SearchBase
    )
    if ($SearchBase[0] -ne ",") {
        $SearchBase = ",{0}" -f $SearchBase
    }
    $ret = $dn -replace $SearchBase, "" -split "(?<!\\)," | ForEach-Object { $_ -split "=" | Select-Object -Skip 1 }
    , $ret
}

function NewOu {
    [CmdletBinding()]
    param (
        [Parameter()]
        [Object]
        $OU,
        $SearchBase
    )
    $ret = New-Object "System.Windows.Controls.TreeViewItem"
    $ret.Header = $OU.Name
    $ret | Add-Member -TypeName "String" -Value $OU.DistinguishedName -MemberType NoteProperty -Name "DistinguishedName"
    $ret | Add-Member -MemberType NoteProperty -Value (SplitDn $OU.DistinguishedName -SearchBase $SearchBase) -Name "Parents"
    $ret
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
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [String]
        $SearchBase
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
        $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase | ForEach-Object { NewOu $_ -SearchBase $SearchBase } | Group-Object { $_.Parents.Count } | Sort-Object { [int]$_.Name }
        $previous = $ous[0].Group | Sort-Object {$_.Header}
        foreach ($i in 1..($ous.Count - 1)) {
            $currentGroup = $ous[$i].Group | Sort-Object {$_.Header}
            $currentGroup | ForEach-Object {
                $CurrentOU = $PSItem
                Write-Debug -Message ("CurrentOU: <{0}> Parents: <{1}>" -f $CurrentOU.Header, ($CurrentOU.Parents | ConvertTo-Json -Compress))
                $previous | Where-Object { $_.Header -EQ $CurrentOU.Parents[1] } | ForEach-Object { 
                    Write-Debug ("Add {0} as child of {1}" -f $CurrentOU.Header, $_.Header )
                    $CurrentOU.Parents = $CurrentOU.Parents | Select-Object -Skip 1
                    $_.AddChild($CurrentOU) 
                }
                $CurrentOU.Group
            }
            $previous = $currentGroup
        }
        $ous[0].Group | Sort-Object {$_.Header} | ForEach-Object { $treeView.Items.Add($_) | Out-Null }

        $DialogResult = $window.ShowDialog()
        if ($DialogResult -eq $true) {
            $treeView.SelectedItem.DistinguishedName
        }
    }

    end {}
}