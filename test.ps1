. .\autoMonitoring.ps1

function condition1($val1, $val2) {
    return $val1 -le $val2;
}

# Configuration
[array]$criticalParams = @(
    @{
        columnIndex = 4;
        curenrValue = 14;
        reqexp      = @{
            re = "[^0-9]";
            converter =  'int'
        }
        condition = (Get-ChildItem "Function:condition1");
    }
);
[array]$validParams = @(
    @{
        rowIndex = 1;
        IndexColumnValueForRowTitle = 1;
        cells = $criticalParams
    }
);

[object]$authForm = $null
<# If not $null
[object]$authForm = @{
    username = Read-Host 'Past username.'
    password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR((Read-Host 'Past user password.' -AsSecureString)));
}
#>

[hashtable]$HTMLTableParams = @{
    id    =  $null;
    tag   =  $null;
    class = 'twc-table';
    
    indexOfClassOrTagElements = 0;
    
    headerAdress = 'th' # 'th' - table head, [int] row number on tbody (0=1);
    firstRow     = 1;
    lastRow      = $null;
};

[string]$aUrl              = $null; # if suit have auth page past it on hear;     
[string]$cUrl              = "https://weather.com/weather/tenday/l/USNY0996:1:US";                                                                                                                                                                                     
[string]$mailTo            = 'my@gmail.com';                                                           
[string]$subject           = 'Weather is Bad!';                                                               
[object]$onError          = [EmailSender]::new($mailTo, $subject);                                                
[int]$onErrorTimeout       = 60;
[Validator]$validator      = [Validator]::new($validParams); 
[MyApplication]$app        = [MyApplication]::new($aUrl, $cUrl, $authForm, $HTMLTableParams, $validator, $onError, $mainTimeout);
$host.ui.RawUI.WindowTitle = "Weather Monitoring";

$app.run();