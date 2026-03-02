Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mass Spectrometry Utility Suite"
        Width="460"
        SizeToContent="Height"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI">
    <StackPanel Margin="16,12,16,16">

        <!-- PROJECT -->
        <DockPanel Margin="0,0,0,4">
            <TextBlock Text="Project" FontWeight="Bold" FontSize="12"
                       DockPanel.Dock="Left" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <Separator VerticalAlignment="Center"/>
        </DockPanel>

        <Button Name="btn1" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <StackPanel>
                <TextBlock Text="[1]  Project folder initializer" FontSize="13"/>
                <TextBlock Text="Creates project structure and logs column info"
                           FontSize="11" Foreground="Gray" Margin="0,2,0,0"/>
            </StackPanel>
        </Button>

        <Button Name="btn2" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <StackPanel>
                <TextBlock Text="[2]  Repair project order" FontSize="13"/>
                <TextBlock Text="Re-numbers projects by creation date and rebuilds column_log.csv"
                           FontSize="11" Foreground="Gray" Margin="0,2,0,0"/>
            </StackPanel>
        </Button>

        <Button Name="btn3" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <StackPanel>
                <TextBlock Text="[3]  Backfill existing column" FontSize="13"/>
                <TextBlock Text="Generates project_info.json and column_log.csv for existing folders"
                           FontSize="11" Foreground="Gray" Margin="0,2,0,0"/>
            </StackPanel>
        </Button>

        <Button Name="btn4" Margin="0,0,0,0" Padding="10,6" HorizontalContentAlignment="Left">
            <StackPanel>
                <TextBlock Text="[4]  Find project" FontSize="13"/>
                <TextBlock Text="Search by ID, project name, or PI"
                           FontSize="11" Foreground="Gray" Margin="0,2,0,0"/>
            </StackPanel>
        </Button>

        <!-- ANALYSIS -->
        <DockPanel Margin="0,12,0,4">
            <TextBlock Text="Analysis" FontWeight="Bold" FontSize="12"
                       DockPanel.Dock="Left" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <Separator VerticalAlignment="Center"/>
        </DockPanel>

        <Button Name="btn5" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <StackPanel>
                <TextBlock Text="[5]  Column usage report" FontSize="13"/>
                <TextBlock Text="All .raw files must be within column parent dir"
                           FontSize="11" Foreground="Gray" Margin="0,2,0,0"/>
            </StackPanel>
        </Button>

        <Button Name="btn6" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <TextBlock Text="[6]  DIA-NN metrics (plots + TSV)" FontSize="13"/>
        </Button>

        <Button Name="btn7" Margin="0,0,0,0" Padding="10,6" HorizontalContentAlignment="Left">
            <StackPanel>
                <TextBlock Text="[7]  Service report (Excel)" FontSize="13"/>
                <TextBlock Text="Analysis_Report.xlsx  (5 sheets)"
                           FontSize="11" Foreground="Gray" Margin="0,2,0,0"/>
            </StackPanel>
        </Button>

        <!-- MISCELLANEOUS -->
        <DockPanel Margin="0,12,0,4">
            <TextBlock Text="Miscellaneous" FontWeight="Bold" FontSize="12"
                       DockPanel.Dock="Left" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <Separator VerticalAlignment="Center"/>
        </DockPanel>

        <Button Name="btn8" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <TextBlock Text="[8]  Bulk convert .raw to mzML  (msConvert)" FontSize="13"/>
        </Button>

        <Button Name="btn9" Margin="0,0,0,3" Padding="10,6" HorizontalContentAlignment="Left">
            <TextBlock Text="[9]  Contaminant check  (mzsniffer)" FontSize="13"/>
        </Button>

        <Button Name="btn99" Margin="0,0,0,0" Padding="10,6" HorizontalContentAlignment="Left">
            <TextBlock Text="[99] Clear method files  (*sld  *meth)" FontSize="13"/>
        </Button>

        <!-- EXIT -->
        <Button Name="btnExit" Margin="0,14,0,0" Padding="10,6"
                HorizontalContentAlignment="Left">
            <TextBlock Text="Exit" FontSize="13"/>
        </Button>

    </StackPanel>
</Window>
'@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

function Invoke-Script ($scriptFile) {
    Start-Process powershell `
        -WorkingDirectory $PSScriptRoot `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ".\$scriptFile"
}

$window.FindName('btn1').Add_Click({   Invoke-Script '1_Project_init.ps1'         })
$window.FindName('btn2').Add_Click({   Invoke-Script '2_Repair_project_order.ps1' })
$window.FindName('btn3').Add_Click({   Invoke-Script '3_Backfill_column.ps1'      })
$window.FindName('btn4').Add_Click({   Invoke-Script '4_Find_project.ps1'         })
$window.FindName('btn5').Add_Click({   Invoke-Script '5_Column_usage.ps1'         })
$window.FindName('btn6').Add_Click({   Invoke-Script '6_DIANN_metrics.ps1'        })
$window.FindName('btn7').Add_Click({   Invoke-Script '7_Report_generator.ps1'     })
$window.FindName('btn8').Add_Click({   Invoke-Script '8_Bulk_msConvert.ps1'       })
$window.FindName('btn9').Add_Click({   Invoke-Script '9_Contaminant_check.ps1'    })
$window.FindName('btn99').Add_Click({  Invoke-Script '99_Clear_files.ps1'         })
$window.FindName('btnExit').Add_Click({ $window.Close() })

[void]$window.ShowDialog()
