<#	
	===========================================================================
	 Created on:   	  	6/24/2016
	 Created by:   	  	TFreedland
	 Last Updated:    	7/29/2016
	 Last Updated by: 	TFreedland
	-------------------------------------------------------------------------
	 Version Number:	1.0.2
	 GUID:			09c0a446-0337-41a5-ac41-fd23fa15ba09
	===========================================================================
#>

#.EXTERNALHELP Format-CSV.psm1-Help.xml
function Format-CSV
{
    [CmdletBinding()]
    [Alias('fcsv')]
    [OutputType('System.Int32')]
    param (
        [Parameter(Mandatory = $true, Position = 0,
                   HelpMessage = '[string] Path to a .csv file containing data to be formatted.')]
        [ValidatePattern('.+\.csv')]
        [string]$InFile,
        
        [Parameter(Mandatory = $true, Position = 1,
                   HelpMessage = '[string] File Path where formatted document should be saved.')]
        [ValidateNotNullOrEmpty()]
        [string]$OutDir,
        
        [Parameter(Mandatory = $false,
                   HelpMessage = '[string] Name of file to be saved in $OutDir.')]
        [string]$SaveName = 'Formatted-Csv',
        
        [switch]$Show
    )
    BEGIN { Set-StrictMode -Version latest }
    PROCESS
    {
        #region prep_data
        [array]$sourceData = Import-Csv $InFile
        [string]$firstLine = $sourceData[0]
        [array]$categories = $firstLine.split(';')
        [array]$headers = @()
        
        # parse csv for column headings
        foreach ($entry in $categories)
        {
            if ($entry -like "@{*") { $entry = $entry.replace('@{', '') }
            if ($entry -like "*}") { $entry = $entry.replace('}', '') }
            if ($entry -match '(.*)=.*') { $entry = "$($matches[1])" }
            $entry = $entry.trim()
            $headers += $entry
        }
        
        [hashtable]$table = @{ }
        # Create one array per header and populate with data from that column.
        $i = 0
        foreach ($item in $headers)
        {
            $table += @{ $item = $sourceData.$item }
            $i++
        }
        #endregion prep_data
        
        #region prep_excel_file_and_output
        $excel = New-Object -Com Excel.Application
        $excel.visible = $false
        $excelWb = $excel.Workbooks.Add()
        $sheetOne = $excelWb.WorkSheets.Item(1)
        $sheetOne.name = 'Data'
        
        # Header Row
        $headerRow = 1
        $headerColumn = 1
        $contentColumn = 1
        $count = ($table.Keys).Count
        $i = 0
        foreach ($key in $table.Keys)
        {
            Write-Progress -Activity "Working On Category: $key" `
                           -Status "#$i Of $count" `
                           -PercentComplete (($i/$count) * 100) `
                           -Id 1
            
            $sheetOne.Cells.Item($headerRow, $headerColumn) = $key
            $headerColumn++
            
            # populate column data
            $contentRow = 2
            $zCount = ($table.$key).Count
            $z = 0
            foreach ($entry in $table.$key)
            {
                Write-Progress -Activity 'Populating Data In Excel' `
                               -Status "Working On #$z Of $zCount" `
                               -PercentComplete (($z/$zCount) * 100) `
                               -Id 2
                
                $sheetOne.Cells.Item($contentRow, $contentColumn) = $entry
                $contentRow++
                $z++
            }
            $contentColumn++
            $i++
        }
        
        $range = $sheetOne.rows(1)
        $range.select()
        $range.Font.ColorIndex = 1
        $range.Font.Bold = $True
        #endregion prep_excel_file_and_output
        
        #region add_table_and_formatting        
        # make tab 1 a table for easier sorting
        $sheetOne.select()
        $sheetOne.ListObjects.Add
        $excel.ActiveSheet.ListObjects.add(1, $excel.ActiveSheet.UsedRange, 0, 1)
        
        # autofit columns and left-center text justification
        $workBook = $sheetOne.UsedRange
        $workBook.Select
        $workBook.EntireColumn.AutoFit()
        $sheetOne.Cells.HorizontalAlignment = -4131 # Left
        $sheetOne.Cells.VerticalAlignment = -4108 # Center
        
        # unhighlight top row
        $range = $sheetOne.Range('A1')
        $range.Select()
        #endregion add_table_and_formatting
        
        #region save_file
        [string]$date = (Get-Date).ToFileTime()
        $SavePath = "$OutDir\$saveName$date.xlsx"
        $sheetOne.SaveAs($SavePath)
        
        # close file (and release file lock)
        $excelWb.Close()
        $excel.Application.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        Remove-Variable Excel
        
        # verify save was successfull
        $checkSave = Test-Path $SavePath
        if ($checkSave -eq $true) { Write-Output "INFO: Report has been successfully saved to: $SavePath" }
        else { Write-Error 'Uh Oh, something went wrong and the excel file could not be saved.' }
        
        if ($Show -eq $true) { Invoke-Item $SavePath }
        #endregion save_file
    }
}

Export-ModuleMember -Function 'Format-CSV' -Alias 'fcsv'
