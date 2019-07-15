class Row
{
    hidden [object]$map      = @{};
    hidden [int]$charsLength = 0;
    
    [int]$rowIndex = $null;

    [object] cells([int]$column)
    {
        return $this.map[$column];
    }

    [string] toString() {
        [string]$curString = '';

        forEach($key in $this.map.getEnumerator() | Sort-Object -Property key)
        {
            $curString += $key.value.toString();
        }

        $curString        = '{0}|' -f $curString;
        $this.charsLength = $curString.Length;

        return $curString;
    }
    
    [void] add($cell)
    {
        $this.map[$cell.columnIndex] = $cell;
    }

    Row([int]$rowIndex)
    {
        $this.rowIndex = $rowIndex;
    }
}


class Cell
{
    [int]$columnIndex    = $null;
    [int]$Row            = $null;
    [string]$columnTitle = $null;
    [object]$value       = $null;
    [Table]$master       = $null;

    hidden [string] optimizeValueToSmartFormat([string]$str)
    {
        [array]$specialList  = ($str -replace "[\s]|\n", "") -split "(.{15})";
        [string]$line        = $specialList[0];
        
        if ($specialList.Count -ne 1) {
            $line = '{0} ...' -f $specialList[1];
        } 
        
        return "| {0}{1}" -f $line, (" " * (20 - $line.Length));
    }

    [object] toString()
    {
        [string]$curResult = $this.optimizeValueToSmartFormat($this.value);

        return $curResult;
    }

    [void] Set([object]$value)
    {
        $this.value = $value;
    }

    [void] del()
    {
        [string]$range = "{0}{1}" -f $this.row, $this.column;
        
        $this.master[$range] = $null;
        $this                = $null;
    }

    Cell([Table]$master, [int]$row, [int]$column)
    {
        $this.Row         = $row;
        $this.columnIndex = $column;
        $this.master      = $master;
    }
}


class Table
{
    [object]$map   = @{};
    [object]$_rows = @{};
    
    
    [object] rows([int]$rowIndex)
    {
        return $this._rows[$rowIndex];
    }

    hidden [string] generateStringFromRowsContent() 
    {
        [string]$curString = '';

        forEach($key in $this._rows.getEnumerator() | Sort-Object -Property key)
        {
            $curString += "{0}`n" -f $key.value.toString();
        }

        return "{0}`n{1}{0}" -f ('-' * $this._rows[1].charsLength), $curString; 
    }

    [string] toString()
    {
        if (!($this._rows[1])) {
            return 'Empty';
        } else {
            return $this.generateStringFromRowsContent();
        }
    }

    [object] cells([int]$row, [int]$column)
    {
        [string]$range = "{0}{1}" -f $row, $column;
        
        return $this.map[$range];
    }

    hidden [void] appendRangeAndRows([int]$row, [int]$column, [string]$columnTitle, [object]$value)
    {
        [string]$range = "{0}{1}" -f $row, $column;
        
        if (!($this._rows[$row]) -or ($this._rows[$row] -eq $null)) {
            $this._rows[$row] = [Row]::new($row);
        }
        
        $this.map[$range]              = [Cell]::new($this, $row, $column);
        $this.map[$range].value        = $value;
        $this.map[$range].columnTitle  = $columnTitle;
        
        $this._rows[$row].add($this.map[$range]);
    }

    [void] add([int]$row, [int]$column, [string]$columnTitle, [object]$value)
    {
        [string]$range = "{0}{1}" -f $row, $column;
        
        if ($this.map[$range] -and ($this.map[$range] -ne $null)) {
            Write-Warning "Range:'{0}' all ready exists!" -f $range;
            exit -b;
        } else {
            $this.appendRangeAndRows($row, $column, $columnTitle, $value);
        }
    }

    Table() {}
}


class HTMLConverter
{
    hidden [string] getCleanValue([object]$node)
    {
        [string]$data = $node.innerHTML;
        

        if ($data -match '<SCRIPT .+?>') {
            $data = $data -Replace '<script[^>]*>[^<]*<\/script[^>]*>', '';
        }
        
        [string]$result = $data -Replace '<.+?>', ' ';

        return $result.trim();
        
    }

    hidden [object] _createShablone([object]$row)
    {
        [array]$curTable   = @();
        [int]$specialIndex = 1;

        forEach ($column in $row)
        {
            [string]$key = $this.getCleanValue($column);

            if ($key -eq "") {
               $key = '{0}' -f $specialIndex;
               $specialIndex++;
            };
         
            $curTable += @{
                'id'    = $key;
                'value' = $null
            }

        }
               
        return $curTable;
    }

    hidden [object] getTdValuesOnRowShablone([object]$cells, [array]$shablone)
    {
        [int]$index  = 0;
        [bool]$empty = $true;

        forEach ($cell in $cells)
        {
            $node  = $shablone[$index];
            $value = $this.getCleanValue($cell);
            
            if (($value -ne $null) -and (($value -replace "[\s]", "") -ne ""))
            {
               $empty = $false; 
            }

            if ($node -ne $null) {
                $node['value'] = $value;
            }

            $index++;
        }
        
        if ($empty) {
            return $null;
        } else        {
            return $shablone;
        }
    }

    hidden [object] shabloneGenerator([object]$HTMLtable,[object]$rows, [object]$headerAdress)
    {
        [object]$row = $null;
        
        if ($headerAdress -eq 'th') {
            $row = $HTMLtable.getElementsByTagName('th');
        } else {
            $row = $rows[$headerAdress].getElementsByTagName('td'); 
        }

        return $this._createShablone($row);
    }

    [Table] convertHTMLTableToObject([object]$HTMLtable, [int]$startRow, [int]$lastRow, [object]$headerAdress)
    {
        [object]$rows       = $HTMLtable.getElementsByTagName("tbody")[0].getElementsByTagName("tr");
        [array]$rowShablone = $this.shabloneGenerator($HTMLtable, $rows, [object]$headerAdress);
        [int]$rowIndex      = 1;
        [int]$curRowIndex   = 1;
        [Table]$curTable    = [Table]::new();
        
        forEach ($row in $rows)
        {
            if ($startRow -le $rowIndex) {
                [object]$cells       = $row.getElementsByTagName("td");
                [object]$curentCells = $this.getTdValuesOnRowShablone($cells, $rowShablone);
                [int]$columnIndex    = 1;

                $curentCells.forEach({
                    $curTable.add($curRowIndex, $columnIndex, $_.id, $_.value);
                    $columnIndex++;
                });

                $curRowIndex ++;
            }

            if ($rowIndex -eq $lastRow) {break;}
            $rowIndex++;
        }

        return $curTable;
    }
}


class Validator
{
    [array]$validParams;
    
    hidden [bool]checking($val, $constant, $condition)
    {
            return !(& $condition($val, $constant));
    }

    hidden [string] __reqexp([string]$value, [object]$reqexpParams)
    {
        [string]$curentValue = $value;

        if ($reqexpParams -ne $null)
        {
            [string]$re        = $reqexpParams.re;
            [object]$converter = $reqexpParams.converter;
            
            if ($re -ne $null) {
                $curentValue = $value -replace $reqexpParams.re, "";
            } 

            if ($converter -ne $null) {
                if ($converter -eq 'int') {
                    $curentValue = [int]::Parse($curentValue, 1);
                } 
                if ($converter -eq 'double') {
                    try {
                        $curentValue = [double]::Parse($curentValue);
                    } catch {
                        $curentValue = [double]::Parse($curentValue.replace(".", ","));
                    }
                }
            }
        }
        
        return $curentValue;
    }
    
    hidden [object] chekRowForError([Row]$row, [object]$rowParams)
    {
        [object]$result = @{
            error       = $false;
            description = ''
        };

        forEach ($column in $rowParams)
        {
            [Cell]$cell    = $row.cells($column.columnIndex);
            $value = $this.__reqexp($cell.value, $column.reqexp);
            [bool]$isError = $this.checking($value, $column.curenrValue, $column.condition);
            
            if ($isError) {
                $result.error        = $true;
                $result.description += "`n{0}: {1}" -f $cell.columnTitle, $value;
            }
        }

        return $result;
    }
     
    [object] checkData([object]$table)
    {  
        [object]$result = @{
            error       = $false;
            description = ''
        };

        for ($n = 0; $n -lt $this.validParams.Count; $n++)
        {
            [array]$param     = $this.validParams[$n];
            [Row]$row         = $table.rows($param.rowIndex)
            [object]$response = $this.chekRowForError($row, $param.cells);
            [string]$rowTitle = $row.cells($param.IndexColumnValueForRowTitle).value;
            
            if ($response.error) {
                 $result.error = $true;
                 $result.description += "{0}{1}`n`n" -f $rowTitle, $response.description;
            }
        }

        return $result;
    }

    Validator ([object]$validParams)
    {
        $this.validParams = $validParams;
    }
}


class MyApplication
{   
    [string]$authUrl            = $null;
    [string]$uri                = $null;
    [object]$authForm           = $null;
    [hashtable]$mainTableParams = $null;
    [object]$onError            = $null;
    [object]$activeError        = $false;
    [bool]$firstAlertWas        = $false;
    [Validator]$validator       = $null;
    [int]$mainTimeout           = $null;
    [HTMLConverter]$converter   = [HTMLConverter]::new();

    hidden [object] __request() 
    {
        [object]$Form       = $this.authForm;
        [object]$my_session = $null;

        if ($this.authUrl) {
            [object]$r    = Invoke-WebRequest -Uri $this.authUrl -SessionVariable my_session -Method POST -Body $Form;
            [object]$data = Invoke-WebRequest -Uri $this.uri     -WebSession $my_session     -Method POST -Body $Form;
        } else {
            [object]$data = Invoke-WebRequest -Uri $this.uri;
        }

        return $data.Content;
    }

    hidden [object] getHTMLContentWithAuthorisation()
    {
        

        try {
           [object]$response = $this.__request();
        } catch {
           [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true};
           [object]$response = $this.__request();
        }

        return $response;
    }

    hidden [object] generateDocument()
    {
        [object]$document = New-Object -ComObject "HTMLFile";
        [object]$content  = $this.getHTMLContentWithAuthorisation();

        $document.IHTMLDocument2_write($content);
        
        return $document;
    }

    hidden [void] __onError__([object]$checkResponse)
    {
        $this.activeError = $checkResponse;
        
        [System.Media.SystemSounds]::Hand.Play();
        $this.onError.main($checkResponse.description);
    }

    hidden __main__([object]$checkResponse, $tbl)
    {
        [bool]$isError = $checkResponse.error;

        if ($isError -and !$this.activeError) {
            if (!$this.firstAlertWas) {
                $this.firstAlertWas = $true;
                sleep($this.mainTimeout);
            } 
            else {
                $this.__onError__($checkResponse);
            }
        } 
        
        if (!$isError) {
            $this.firstAlertWas = $false;
            $this.activeError   = $false;
        }

        $this.run();
    }

    hidden [object] forTest()
    {
        
        [object]$html        = New-Object -ComObject "HTMLFile";
        [object]$content     = Get-Content -Path ".\tableTest.html"  -Raw;
        [HTMLConverter]$pars = [HTMLConverter]::new();
        
        $html.IHTMLDocument2_write($content);
        $htmltable = $html.getElementById($this.mainTableId);
        
        return $pars.convertHTMLTableToObject($htmltable);
    }

    hidden [object] getHTMLTable([object]$document)
    {
        [int]$index     = $this.mainTableParams.indexOfClassOrTagElements;
        [string]$id     = $this.mainTableParams.id;
        [string]$tag    = $this.mainTableParams.tag;
        [string]$class  = $this.mainTableParams.class;
        [object]$result = $null;
       
        if ($index -eq $null) {$index = 0;}

        if ($id -ne $null)  {
            $result = $document.getElementById($id);
        } 
        
        if ($tag -ne $null) {
           $result = $document.getElementsByTagName($tag)[$index];
        }
        
        if ($class -ne $null) {
            $result = $document.getElementsByClassName($class)[$index];
        }

        return $result;
    }

    [void] run()
    {
        #[object]$mainTable = $this.forTest();
        
        [object]$document  = $this.generateDocument();
        [object]$HTMLtabel = $this.getHTMLTable($document);
        [object]$mainTable = $this.converter.convertHTMLTableToObject($HTMLtabel, $this.mainTableParams.firstRow, $this.mainTableParams.lastRow, $this.mainTableParams.headerAdress);
        [object]$check     = $this.validator.checkData($mainTable);
        
        clear;
        write-host ($mainTable.toString())`n($this.uri)`n($this.activeError.description);
        sleep(60);
        
        $this.__main__($check, $mainTable);
    }

    MyApplication([string]$authUrl, [string]$uri, [object]$authForm, [hashtable]$mainTableParams, [Validator]$validator, [object]$onError, [int]$mainTimeout)
    {
        $this.authUrl         = $authUrl;
        $this.uri             = $uri;
        $this.authForm        = $authForm;
        $this.mainTableParams = $mainTableParams;
        $this.validator       = $validator;
        $this.onError         = $onError;
        $this.mainTimeout     = $mainTimeout;
    }
}


class EmailSender
{
    [string]$emailTo;
    [string]$mailSubject;

    [void] main([string]$data)
    {
        [object]$outlook = New-Object -ComObject Outlook.Application;
        [object]$mail    = $outlook.CreateItem(0);
        
        $mail.To      = $this.emailTo;
        $mail.Subject = $this.mailSubject; 
        $mail.Body    = $data;

        $mail.send();
    }

    EmailSender([string]$emailTo, [string]$mailSubject)
    {
        $this.emailTo     = $emailTo;
        $this.mailSubject = $mailSubject;
    }
}