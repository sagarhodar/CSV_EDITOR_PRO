# ================================================
# Script: Edit-CsvMatrix.ps1
# Purpose: Load CSV into editable grid GUI (no Excel)
# Features: Edit, Add/Delete Rows/Columns, Column Filters (>, <, >=, <=, =), Global Search, Save, Save As
# ================================================

param (
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = ""
)

# Check file existence
if ($CsvPath -ne "" -and !(Test-Path $CsvPath)) {
    [System.Windows.Forms.MessageBox]::Show("File not found: $CsvPath", "Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Create main form
$form = New-Object System.Windows.Forms.Form
if ($CsvPath -ne "") {
    $form.Text = "CSV Editor - $([System.IO.Path]::GetFileName($CsvPath))"
} else {
    $form.Text = "CSV Editor - [New File]"
}
$form.Size = New-Object System.Drawing.Size(1100, 800)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Global Search Panel ---
$searchPanel = New-Object System.Windows.Forms.Panel
$searchPanel.Location = New-Object System.Drawing.Point(10, 10)
$searchPanel.Size = New-Object System.Drawing.Size(1060, 35)
$searchPanel.Anchor = 'Top,Left,Right'

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search All:"
$searchLabel.Location = New-Object System.Drawing.Point(0, 10)
$searchLabel.Size = New-Object System.Drawing.Size(70, 20)

$globalSearchBox = New-Object System.Windows.Forms.TextBox
$globalSearchBox.Location = New-Object System.Drawing.Point(75, 7)
$globalSearchBox.Size = New-Object System.Drawing.Size(300, 20)
$globalSearchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$searchInfo = New-Object System.Windows.Forms.Label
$searchInfo.Text = "Live search across all columns and rows"
$searchInfo.Location = New-Object System.Drawing.Point(385, 10)
$searchInfo.Size = New-Object System.Drawing.Size(300, 20)
$searchInfo.ForeColor = [System.Drawing.Color]::Gray

$searchPanel.Controls.AddRange(@($searchLabel, $globalSearchBox, $searchInfo))

# --- Column Management Panel ---
$columnPanel = New-Object System.Windows.Forms.Panel
$columnPanel.Location = New-Object System.Drawing.Point(10, 50)
$columnPanel.Size = New-Object System.Drawing.Size(1060, 35)
$columnPanel.Anchor = 'Top,Left,Right'

$addColButton = New-Object System.Windows.Forms.Button
$addColButton.Text = "[+] Add Column"
$addColButton.Location = New-Object System.Drawing.Point(0, 5)
$addColButton.Size = New-Object System.Drawing.Size(110, 25)

$deleteColButton = New-Object System.Windows.Forms.Button
$deleteColButton.Text = "[-] Delete Column"
$deleteColButton.Location = New-Object System.Drawing.Point(120, 5)
$deleteColButton.Size = New-Object System.Drawing.Size(120, 25)

$colLabel = New-Object System.Windows.Forms.Label
$colLabel.Text = "Inserts next to selected column | Column filters support: >5, <10, >=20, <=100, =exact, or text"
$colLabel.Location = New-Object System.Drawing.Point(250, 10)
$colLabel.Size = New-Object System.Drawing.Size(600, 20)
$colLabel.ForeColor = [System.Drawing.Color]::Gray

$columnPanel.Controls.AddRange(@($addColButton, $deleteColButton, $colLabel))

# --- DataGridView ---
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 130)
$grid.Size = New-Object System.Drawing.Size(1060, 520)
$grid.Anchor = 'Top,Left,Right,Bottom'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = 'Fill'
$grid.RowHeadersWidth = 30
$grid.ReadOnly = $false
$grid.EditMode = 'EditOnKeystrokeOrF2'
$grid.ColumnHeadersHeight = 28
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray

# Load CSV data
$table = New-Object System.Data.DataTable

if ($CsvPath -ne "" -and (Test-Path $CsvPath)) {
    # Load from existing file
    $data = Import-Csv -Path $CsvPath
    
    if ($data.Count -gt 0) {
        # Convert to DataTable
        $headers = $data[0].PSObject.Properties.Name
        foreach ($h in $headers) { [void]$table.Columns.Add($h) }

        foreach ($row in $data) {
            $dr = $table.NewRow()
            foreach ($h in $headers) { $dr[$h] = $row.$h }
            $table.Rows.Add($dr)
        }
    } else {
        # Handle empty CSV - try to get headers
        $firstLine = Get-Content -Path $CsvPath -TotalCount 1
        if ($firstLine) {
            $headers = $firstLine -split ','
            foreach ($h in $headers) { 
                [void]$table.Columns.Add($h.Trim().Trim('"')) 
            }
        }
    }
} else {
    # Create new blank CSV
    [void]$table.Columns.Add("Column1")
    [void]$table.Columns.Add("Column2")
    [void]$table.Columns.Add("Column3")
    
    # Add one empty row
    $newRow = $table.NewRow()
    $table.Rows.Add($newRow)
}

# Create BindingSource for filtering
$bindingSource = New-Object System.Windows.Forms.BindingSource
$bindingSource.DataSource = $table
$grid.DataSource = $bindingSource

# Hashtable to hold filter TextBoxes
$filterBoxes = @{}

# --- Column Filter Functions ---
function Parse-FilterExpression {
    param([string]$text, [string]$colName)
    
    $text = $text.Trim()
    if ($text -eq "") { return "" }
    
    # Check for comparison operators: >=, <=, >, <, =
    if ($text -match '^(>=|<=|>|<|=)\s*(.+)$') {
        $op = $matches[1]
        $val = $matches[2].Trim()
        $colEsc = $colName.Replace("]", "]]")
        
        # Try to determine if it's numeric
        $numVal = 0
        $isNumeric = [double]::TryParse($val, [ref]$numVal)
        
        if ($isNumeric) {
            # Numeric comparison
            switch ($op) {
                ">=" { return "CONVERT([$colEsc], 'System.Double') >= $numVal" }
                "<=" { return "CONVERT([$colEsc], 'System.Double') <= $numVal" }
                ">"  { return "CONVERT([$colEsc], 'System.Double') > $numVal" }
                "<"  { return "CONVERT([$colEsc], 'System.Double') < $numVal" }
                "="  { return "CONVERT([$colEsc], 'System.Double') = $numVal" }
            }
        } else {
            # String comparison
            $valEsc = $val.Replace("'", "''")
            switch ($op) {
                ">=" { return "[$colEsc] >= '$valEsc'" }
                "<=" { return "[$colEsc] <= '$valEsc'" }
                ">"  { return "[$colEsc] > '$valEsc'" }
                "<"  { return "[$colEsc] < '$valEsc'" }
                "="  { return "[$colEsc] = '$valEsc'" }
            }
        }
    }
    
    # Default: LIKE search (contains)
    $colEsc = $colName.Replace("]", "]]")
    $searchText = $text.Replace("'", "''")
    return "CONVERT([$colEsc], 'System.String') LIKE '%$searchText%'"
}

function Update-ColumnFilters {
    $filters = @()
    
    # Add global search filter if present
    $globalText = $globalSearchBox.Text.Trim()
    if ($globalText -ne "") {
        $globalFilters = @()
        foreach ($col in $table.Columns) {
            $colEsc = $col.ColumnName.Replace("]", "]]")
            $searchText = $globalText.Replace("'", "''")
            $globalFilters += "CONVERT([$colEsc], 'System.String') LIKE '%$searchText%'"
        }
        if ($globalFilters.Count -gt 0) {
            $filters += "(" + ([string]::Join(" OR ", $globalFilters)) + ")"
        }
    }
    
    # Add column-specific filters
    foreach ($colName in $filterBoxes.Keys) {
        $filterText = $filterBoxes[$colName].Text.Trim()
        if ($filterText -ne "") {
            $expr = Parse-FilterExpression -text $filterText -colName $colName
            if ($expr -ne "") {
                $filters += "($expr)"
            }
        }
    }
    
    try {
        if ($filters.Count -gt 0) {
            $bindingSource.Filter = [string]::Join(" AND ", $filters)
        } else {
            $bindingSource.Filter = ""
        }
    } catch {
        # If filter fails, clear it
        $bindingSource.Filter = ""
    }
}

function Create-ColumnFilters {
    # Remove old filters
    foreach ($tb in $filterBoxes.Values) {
        if ($form.Controls.Contains($tb)) { $form.Controls.Remove($tb) }
        try { $tb.Dispose() } catch {}
    }
    $filterBoxes.Clear()
    
    # Create filter textbox for each column
    for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
        $colName = $grid.Columns[$i].Name
        
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Name = "Filter_$colName"
        $tb.Tag = $colName
        $tb.Height = 22
        $tb.BorderStyle = 'FixedSingle'
        $tb.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        
        # Tooltip
        $tt = New-Object System.Windows.Forms.ToolTip
        $tt.SetToolTip($tb, "$colName`nSupports: >5, <10, >=20, <=100, =exact, or text search")
        
        $form.Controls.Add($tb)
        $filterBoxes[$colName] = $tb
        
        # TextChanged event
        $tb.Add_TextChanged({
            Update-ColumnFilters
        })
    }
    
    Position-FilterBoxes
}

function Position-FilterBoxes {
    for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
        $colName = $grid.Columns[$i].Name
        if (-not $filterBoxes.ContainsKey($colName)) { continue }
        
        $tb = $filterBoxes[$colName]
        
        try {
            $rect = $grid.GetCellDisplayRectangle($i, -1, $false)
            if ($rect.Width -le 4 -or $rect.X -lt 0) {
                $tb.Visible = $false
                continue
            }
            
            $tb.Visible = $true
            $tb.Width = [Math]::Max(50, $rect.Width - 6)
            $tb.Left = $grid.Left + $rect.X + 3
            $tb.Top = $grid.Top - $tb.Height - 6
        } catch {
            $tb.Visible = $false
        }
    }
}

# Grid events for filter positioning
$grid.add_ColumnWidthChanged({ Position-FilterBoxes })
$grid.add_Scroll({ Position-FilterBoxes })
$grid.add_SizeChanged({ Position-FilterBoxes })
$grid.add_ColumnDisplayIndexChanged({ Position-FilterBoxes })

# Global search box event
$globalSearchBox.Add_TextChanged({
    Update-ColumnFilters
})

# --- Add Column (Next to Selected) ---
$addColButton.Add_Click({
    $colName = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter new column name:", 
        "Add Column", 
        "NewColumn")
    
    if ($colName -ne "" -and $colName -ne $null) {
        try {
            if ($table.Columns.Contains($colName)) {
                [System.Windows.Forms.MessageBox]::Show("Column '$colName' already exists.", "Duplicate Column")
                return
            }
            
            # Get selected column index
            $insertIndex = -1
            if ($grid.CurrentCell -ne $null) {
                $insertIndex = $grid.CurrentCell.ColumnIndex + 1
            }
            
            # Add column first, then reorder
            $newCol = New-Object System.Data.DataColumn($colName)
            $table.Columns.Add($newCol)
            
            # Set position if we have a valid insert index
            if ($insertIndex -ge 0 -and $insertIndex -lt $table.Columns.Count) {
                $newCol.SetOrdinal($insertIndex)
            }
            
            $bindingSource.ResetBindings($false)
            
            # Recreate filters
            Create-ColumnFilters
            
            [System.Windows.Forms.MessageBox]::Show("Column '$colName' added successfully.", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error adding column: $_", "Error")
        }
    }
})

# --- Delete Column ---
$deleteColButton.Add_Click({
    if ($grid.CurrentCell -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please click on a column to select it first.", "No Column Selected")
        return
    }
    
    $colIndex = $grid.CurrentCell.ColumnIndex
    $colName = $grid.Columns[$colIndex].Name
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Delete column '$colName'?`nThis cannot be undone.", 
        "Confirm Delete Column", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $table.Columns.Remove($colName)
            $bindingSource.ResetBindings($false)
            
            # Recreate filters
            Create-ColumnFilters
            
            [System.Windows.Forms.MessageBox]::Show("Column '$colName' deleted successfully.", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error deleting column: $_", "Error")
        }
    }
})

# --- Button Panel ---
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10, 660)
$buttonPanel.Size = New-Object System.Drawing.Size(1060, 70)
$buttonPanel.Anchor = 'Bottom,Left,Right'

# Open File Button
$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = "Open CSV..."
$openButton.Location = New-Object System.Drawing.Point(0, 5)
$openButton.Size = New-Object System.Drawing.Size(90, 30)
$openButton.BackColor = [System.Drawing.Color]::LightBlue

$openButton.Add_Click({
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openDialog.Title = "Open CSV File"
    
    if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:CsvPath = $openDialog.FileName
            
            # Reload data
            $data = Import-Csv -Path $CsvPath
            $table.Clear()
            $table.Columns.Clear()
            
            if ($data.Count -gt 0) {
                $headers = $data[0].PSObject.Properties.Name
                foreach ($h in $headers) { [void]$table.Columns.Add($h) }
                
                foreach ($row in $data) {
                    $dr = $table.NewRow()
                    foreach ($h in $headers) { $dr[$h] = $row.$h }
                    $table.Rows.Add($dr)
                }
            }
            
            $bindingSource.ResetBindings($false)
            Create-ColumnFilters
            $form.Text = "CSV Editor - $([System.IO.Path]::GetFileName($CsvPath))"
            
            [System.Windows.Forms.MessageBox]::Show("File opened successfully.", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening file:`n$_", "Error")
        }
    }
})

# New File Button
$newButton = New-Object System.Windows.Forms.Button
$newButton.Text = "New CSV"
$newButton.Location = New-Object System.Drawing.Point(100, 5)
$newButton.Size = New-Object System.Drawing.Size(90, 30)
$newButton.BackColor = [System.Drawing.Color]::LightGreen

$newButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Create a new blank CSV? Any unsaved changes will be lost.", 
        "New CSV", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            # Clear existing data
            $table.Clear()
            $table.Columns.Clear()
            
            # Add default columns
            [void]$table.Columns.Add("Column1")
            [void]$table.Columns.Add("Column2")
            [void]$table.Columns.Add("Column3")
            
            # Add one empty row
            $newRow = $table.NewRow()
            $table.Rows.Add($newRow)
            
            $bindingSource.ResetBindings($false)
            Create-ColumnFilters
            $form.Text = "CSV Editor - [New File]"
            $script:CsvPath = ""
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating new file:`n$_", "Error")
        }
    }
})

# Add Row Button (Insert Next to Selected)
$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "[+] Add Row"
$addButton.Location = New-Object System.Drawing.Point(200, 5)
$addButton.Size = New-Object System.Drawing.Size(100, 30)
$addButton.BackColor = [System.Drawing.Color]::LightGreen

$addButton.Add_Click({
    try {
        $newRow = $table.NewRow()
        
        # Get selected row index
        $insertIndex = -1
        if ($grid.SelectedRows.Count -gt 0) {
            $insertIndex = $grid.SelectedRows[0].Index + 1
        } else {
            # If no selection, add at end
            $insertIndex = $table.Rows.Count
        }
        
        # Insert at position
        if ($insertIndex -ge 0 -and $insertIndex -lt $table.Rows.Count) {
            $table.Rows.InsertAt($newRow, $insertIndex)
        } else {
            $table.Rows.Add($newRow)
            $insertIndex = $table.Rows.Count - 1
        }
        
        $bindingSource.ResetBindings($false)
        
        # Select the new row
        if ($insertIndex -ge 0 -and $insertIndex -lt $grid.Rows.Count) {
            $grid.ClearSelection()
            $grid.Rows[$insertIndex].Selected = $true
            $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $insertIndex - 5)
            $grid.CurrentCell = $grid.Rows[$insertIndex].Cells[0]
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error adding row: $_", "Error")
    }
})

# Delete Row Button
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "[-] Delete Row(s)"
$deleteButton.Location = New-Object System.Drawing.Point(310, 5)
$deleteButton.Size = New-Object System.Drawing.Size(120, 30)
$deleteButton.BackColor = [System.Drawing.Color]::LightCoral

$deleteButton.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select rows to delete.", "No Selection")
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Delete $($grid.SelectedRows.Count) selected row(s)?", 
        "Confirm Delete", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $rowsToDelete = @()
        foreach ($row in $grid.SelectedRows) {
            $rowsToDelete += $row
        }
        $rowsToDelete = $rowsToDelete | Sort-Object { $_.Index } -Descending
        
        foreach ($row in $rowsToDelete) {
            if ($row.DataBoundItem -is [System.Data.DataRowView]) {
                $row.DataBoundItem.Row.Delete()
            }
        }
        
        $table.AcceptChanges()
        $bindingSource.ResetBindings($false)
    }
})

# Save Button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Location = New-Object System.Drawing.Point(440, 5)
$saveButton.Size = New-Object System.Drawing.Size(90, 30)
$saveButton.BackColor = [System.Drawing.Color]::LightSkyBlue

$saveButton.Add_Click({
    try {
        # Check if we have a file path
        if ($CsvPath -eq "" -or $CsvPath -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Please use 'Save As...' to specify a file name first.", "No File Path")
            return
        }
        
        $sw = New-Object System.IO.StreamWriter($CsvPath, $false, [System.Text.Encoding]::UTF8)
        try {
            # Write header
            $headers = $table.Columns | ForEach-Object { $_.ColumnName }
            $headerLine = ($headers | ForEach-Object { "`"$($_ -replace '"', '""')`"" }) -join ','
            $sw.WriteLine($headerLine)
            
            # Write data rows
            foreach ($row in $table.Rows) {
                if ($row.RowState -ne [System.Data.DataRowState]::Deleted) {
                    $values = foreach ($col in $headers) {
                        $val = $row[$col]
                        if ($val -eq $null) { "" }
                        else { $val.ToString() -replace '"', '""' }
                    }
                    $dataLine = ($values | ForEach-Object { "`"$_`"" }) -join ','
                    $sw.WriteLine($dataLine)
                }
            }
        } finally {
            $sw.Close()
        }
        
        [System.Windows.Forms.MessageBox]::Show("Saved successfully to:`n$CsvPath", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving file:`n$_", "Error")
    }
})

# Save As Button
$saveAsButton = New-Object System.Windows.Forms.Button
$saveAsButton.Text = "Save As..."
$saveAsButton.Location = New-Object System.Drawing.Point(540, 5)
$saveAsButton.Size = New-Object System.Drawing.Size(100, 30)
$saveAsButton.BackColor = [System.Drawing.Color]::LightYellow

$saveAsButton.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveDialog.FileName = [System.IO.Path]::GetFileName($CsvPath)
    $saveDialog.Title = "Save CSV As..."
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $outputPath = $saveDialog.FileName
            
            # Determine which data to save
            $dataToSave = $table
            if ($bindingSource.Filter -ne "") {
                $dataToSave = $table.DefaultView.ToTable()
            }
            
            $sw = New-Object System.IO.StreamWriter($outputPath, $false, [System.Text.Encoding]::UTF8)
            try {
                # Write header
                $headers = $dataToSave.Columns | ForEach-Object { $_.ColumnName }
                $headerLine = ($headers | ForEach-Object { "`"$($_ -replace '"', '""')`"" }) -join ','
                $sw.WriteLine($headerLine)
                
                # Write data rows
                foreach ($row in $dataToSave.Rows) {
                    if ($row.RowState -ne [System.Data.DataRowState]::Deleted) {
                        $values = foreach ($col in $headers) {
                            $val = $row[$col]
                            if ($val -eq $null) { "" }
                            else { $val.ToString() -replace '"', '""' }
                        }
                        $dataLine = ($values | ForEach-Object { "`"$_`"" }) -join ','
                        $sw.WriteLine($dataLine)
                    }
                }
            } finally {
                $sw.Close()
            }
            
            # Update current file path
            $script:CsvPath = $outputPath
            $form.Text = "CSV Editor - $([System.IO.Path]::GetFileName($CsvPath))"
            
            [System.Windows.Forms.MessageBox]::Show("Saved successfully to:`n$outputPath", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving file:`n$_", "Error")
        }
    }
})

# Help Button
$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Text = "Help"
$helpButton.Location = New-Object System.Drawing.Point(650, 5)
$helpButton.Size = New-Object System.Drawing.Size(80, 30)
$helpButton.BackColor = [System.Drawing.Color]::LightGray

$helpButton.Add_Click({
    $helpText = @"
CSV MATRIX EDITOR
=================

FEATURES:
---------
• Edit cells directly (click and type)
• Global Search: Search across all columns
• Column Filters: Filter individual columns
  - Supports: >5, <10, >=20, <=100, =exact, or text search
• Add/Delete Rows: Insert next to selected row
• Add/Delete Columns: Insert next to selected column
• Open CSV: Load existing CSV files
• New CSV: Create blank CSV file
• Save: Save to current file
• Save As: Save with new name (exports filtered data)

KEYBOARD SHORTCUTS:
-------------------
• F2 or Click: Edit cell
• Ctrl+Click: Select multiple rows
• Delete: Remove selected content

TIPS:
-----
• Rows/Columns are inserted next to selected item
• Use column filters for precise filtering
• Save As exports only filtered rows if filter is active
• All data is saved as UTF-8 encoded CSV

=================
Created by: Sagar Hodar
Version: 2.0
=================
"@
    
    [System.Windows.Forms.MessageBox]::Show($helpText, "Help - CSV Matrix Editor", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(0, 40)
$statusLabel.Size = New-Object System.Drawing.Size(1000, 20)
$statusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$buttonPanel.Controls.AddRange(@($openButton, $newButton, $addButton, $deleteButton, $saveButton, $saveAsButton, $helpButton, $statusLabel))

# --- Add all controls to form ---
$form.Controls.AddRange(@($searchPanel, $columnPanel, $grid, $buttonPanel))

# --- Status bar update ---
$updateStatus = {
    $totalRows = $table.Rows.Count
    $visibleRows = $grid.Rows.Count
    if ($bindingSource.Filter -ne "") {
        $statusLabel.Text = "Showing $visibleRows of $totalRows rows (filtered) | Rows/Columns insert next to selected | Click cells to edit"
    } else {
        $statusLabel.Text = "Total rows: $totalRows | Total columns: $($table.Columns.Count) | Rows/Columns insert next to selected | Click cells to edit"
    }
}

$grid.Add_DataBindingComplete($updateStatus)

# Create filters when form is shown (fixes BeginInvoke error)
$form.Add_Shown({
    Create-ColumnFilters
    & $updateStatus
})

# Show GUI
[void]$form.ShowDialog()
