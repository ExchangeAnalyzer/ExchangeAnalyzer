#This function rolls up test results into an object
Function Get-TestResultObject()
{
    [CmdletBinding()]
    param (
	    [Parameter()]
	    $ExchangeAnalyzerTests,
        
        [Parameter()]
	    $TestID,

	    [Parameter()]
	    $PassedList,

	    [Parameter()]
	    $FailedList,

	    [Parameter()]
	    $WarningList,

	    [Parameter()]
	    $InfoList,

	    [Parameter()]
	    $ErrorList
	)

    Write-Verbose "Rolling test result object for $TestID"
    
    $TestComments = $null

    $ReferenceURLBase = "https://github.com/cunninghamp/ExchangeAnalyzer/wiki"
    
    #A test can only pass if there are items in $passedlist and no failed, warning, or info items 
    if ($PassedList -and -not $FailedList -and -not $WarningList -and -not $InfoList)
    {
        $TestComments = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).IfPassedComments
        $TestOutcome = "Passed"
    }

    #A test that has a possibility of an Info outcome should not have any possibility of failed
    #or warning items, but can still throw a warning if the test fails to run or encounters other
    #errors while running.
    if ($InfoList)
    {
        $TestComments = ($ExchangeAnalyzerTests.Test | Where {$_.id -eq $TestID}).IfInfoComments
        $TestOutcome = "Info"
    }

    #If any items are in $warninglist the overall outcome is Warning, unless failed items are
    #encountered next
    if ($WarningList)
    {
        $TestComments = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).IfWarningComments
        $TestOutcome = "Warning"
    }

    #If any items are in $failedlist the overall outcome is Failed
    if ($FailedList)
    {
        $TestComments = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).IfFailedComments
        $TestOutcome = "Failed"
    }

    #If no passed, failed, warning, or info items exist the test likely had an unexpected error.
    if (-not $PassedList -and -not $FailedList -and -not $WarningList -and -not $InfoList)
    {
        $TestComments = "Test could not run."
        $TestOutcome = "Warning"
    }

    if ($ErrorList)
    {
        $TestComments = "Errors were encountered. $($ErrorList)"
        $TestOutcome = "Warning"
    }

    $result = [Ordered]@{
        TestID = $TestID
        TestCategory = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).Category
        TestName = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).Name
        TestDescription = ($ExchangeAnalyzerTests.Test | Where {$_.Id -eq $TestID}).Description
        TestOutcome = $TestOutcome
        PassedObjects = $PassedList
        FailedObjects = $FailedList
        WarningObjects = $WarningList
        InfoObjects = $InfoList
        Comments = $TestComments
        Reference = "$($ReferenceURLBase)/$($TestID)"
    }
    
    $TestResultObj = New-Object -TypeName PSObject -Property $result 
 
    return $TestResultObj
}

#This function scrapes the TechNet page for Exchange Server build numbers and release
#dates to match the build numbers for Exchange Servers in the organization.
#Reference: Lee Holmes article on extracting tables from web pages was very useful for developing this
#Link: #http://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/
# 30/3/2016 - This function has been removed from use by the tests but will remain in the module for now.
Function Get-ExchangeBuildNumbers()
{
    [CmdletBinding()]
    param ()
    
    Write-Verbose "Fetching Exchange build numbers from TechNet"
    
    $URL = "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.160).aspx"
    try
    {
        $WebPage = Invoke-WebRequest -Uri $URL -ErrorAction STOP
        $tables = @($WebPage.Parsedhtml.getElementsByTagName("TABLE"))

        Write-Verbose "Parsing results from web request"
        foreach ($table in $tables)
        {
            $rows = @($table.Rows)

            foreach($row in $rows)
            {
                $cells = @($row.Cells)

                ## If we’ve found a table header, remember its titles
                if($cells[0].tagName -eq "TH")
                {
                    $titles = @($cells | ForEach-Object { ("" + $_.InnerText).Trim() })
                    continue
                }

                ## If we haven’t found any table headers, make up names "P1", "P2", etc.
                if(-not $titles)
                {
                    $titles = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
                }

                ## Now go through the cells in the the row. For each, try to find the
                ## title that represents that column and create a hashtable mapping those
                ## titles to content

                $resultObject = [Ordered] @{}

                for($counter = 0; $counter -lt $cells.Count; $counter++)
                {
                    $title = $titles[$counter]
                    if(-not $title) { continue }
                    $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
                }

                ## And finally cast that hashtable to a PSCustomObject
                [PSCustomObject] $resultObject
            }
    }
    }
    catch
    {
        Write-Warning $_.Exception.Message
        $ExchangeBuildNumbers = "An error occurred. $($_.Exception.Message)"
    }

    return $ExchangeBuildNumbers
}

#This function retrieves the Client Access URLs for HTTPS services
Function Get-ExchangeURLs()
{
    [CmdletBinding()]
    param (
        [Parameter()]
	    $ClientAccessServers
    )

    $results = @()

    foreach ($CAS in $ClientAccessServers)
    {
        Write-Verbose "Fetching URLs for $CAS"
        
        #Outlook Anywhere
        $OA = Get-OutlookAnywhere -Server $CAS.Name -AdPropertiesOnly | Select InternalHostName,ExternalHostName
        if ($OA.InternalHostname -eq $null) { $OA.InternalHostName = "Not set" }
        if ($OA.ExternalHostname -eq $null) { $OA.ExternalHostName = "Not set" }
        
        #Outlook on the web
        $OWA = Get-OWAVirtualDirectory -Server $CAS.Name -AdPropertiesOnly | Select InternalURL,ExternalURL
        if ($OWA.InternalURL -eq $null)
        {
            $OWA.InternalURL = "Not set"
        }
        else
        {
            $OWA.InternalURL = $OWA.InternalURL.AbsoluteUri
        }
        if ($OWA.ExternalURL -eq $null)
        {
            $OWA.ExternalURL = "Not set"
        }
        else
        {
            $OWA.ExternalURL = $OWA.ExternalURL.AbsoluteUri
        }
                
        #Exchange Control Panel
        $ECP = Get-ECPVirtualDirectory -Server $CAS.Name -AdPropertiesOnly | Select InternalURL,ExternalURL
        if ($ECP.InternalURL -eq $null)
        {
            $ECP.InternalURL = "Not set"
        }
        else
        {
            $ECP.InternalURL = $ECP.InternalURL.AbsoluteUri
        }
        if ($ECP.ExternalURL -eq $null)
        {
            $ECP.ExternalURL = "Not set"
        }
        else
        {
            $ECP.ExternalURL = $ECP.ExternalURL.AbsoluteUri
        }
               
        #Offline Address Book        
        $OAB = Get-OABVirtualDirectory -Server $CAS.Name -AdPropertiesOnly | Select InternalURL,ExternalURL
        if ($OAB.InternalURL -eq $null)
        {
            $OAB.InternalURL = "Not set"
        }
        else
        {
            $OAB.InternalURL = $OAB.InternalURL.AbsoluteUri
        }
        if ($OAB.ExternalURL -eq $null)
        {
            $OAB.ExternalURL = "Not set"
        }
        else
        {
            $OAB.ExternalURL = $OAB.ExternalURL.AbsoluteUri
        }
                
        #Exchange Web Services
        $EWS = Get-WebServicesVirtualDirectory -Server $CAS.Name -AdPropertiesOnly | Select InternalURL,ExternalURL
        if ($EWS.InternalURL -eq $null)
        {
            $EWS.InternalURL = "Not set"
        }
        else
        {
            $EWS.InternalURL = $EWS.InternalURL.AbsoluteUri
        }
        if ($EWS.ExternalURL -eq $null)
        {
            $EWS.ExternalURL = "Not set"
        }
        else
        {
            $EWS.ExternalURL = $EWS.ExternalURL.AbsoluteUri
        }
                
        #MAPI
        $MAPI = Get-MAPIVirtualDirectory -Server $CAS.Name -AdPropertiesOnly | Select InternalURL,ExternalURL
        if ($MAPI.InternalURL -eq $null)
        {
            $MAPI.InternalURL = "Not set"
        }
        else
        {
            $MAPI.InternalURL = $MAPI.InternalURL.AbsoluteUri
        }
        if ($MAPI.ExternalURL -eq $null)
        {
            $MAPI.ExternalURL = "Not set"
        }
        else
        {
            $MAPI.ExternalURL = $MAPI.ExternalURL.AbsoluteUri
        }
               
        #ActiveSync
        $EAS = Get-ActiveSyncVirtualDirectory -Server $CAS.Name -AdPropertiesOnly | Select InternalURL,ExternalURL
        if ($EAS.InternalURL -eq $null)
        {
            $EAS.InternalURL = "Not set"
        }
        else
        {
            $EAS.InternalURL = $EAS.InternalURL.AbsoluteUri
        }
        if ($EAS.ExternalURL -eq $null)
        {
            $EAS.ExternalURL = "Not set"
        }
        else
        {
            $EAS.ExternalURL = $EAS.ExternalURL.AbsoluteUri
        }
                
        #AutoDiscover
        Write-Verbose "Fetching AutoD SCP for $CAS"
        $AutoD = Get-ClientAccessServer $CAS.Name -WarningAction Ignore | Select AutoDiscoverServiceInternalUri
        if ($AutoD.AutoDiscoverServiceInternalUri -eq $null) { $AutoD.AutoDiscoverServiceInternalUri -eq "Not set" }

        Write-Verbose "Creating object for CAS Urls"
        $props = [Ordered]@{
            Name = $CAS.Name
            OAInternal = $OA.InternalHostName
            OAExternal = $OA.ExternalHostName
            OWAInternal = $OWA.InternalURL
            OWAExternal = $OWA.ExternalURL
            ECPInternal = $ECP.InternalURL
            ECPExternal = $ECP.ExternalURL
            OABInternal = $OAB.InternalURL
            OABExternal = $OAB.ExternalURL
            EWSInternal = $EWS.InternalURL
            EWSExternal = $EWS.ExternalURL
            MAPIInternal = $MAPI.InternalURL
            MAPIExternal = $MAPI.ExternalURL
            EASInternal = $EAS.InternalURL
            EASExternal = $EAS.ExternalURL
            AutoDSCP = $AutoD.AutoDiscoverServiceInternalUri
            }
            
        $CASObj = New-Object -TypeName PSObject -Property $props

        $results += $CASObj
    }

    return $results
}

#this function reads a property out of the global server property bag
Function Get-ExAServerProperty()
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Server,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $Property
    )

    try
    { 
        Write-Debug "Searching for property '$Property' for server '$Server'"

        if ($ExAPropertyBag -eq $null)
        {
            throw 'ExAPropertyBag cannot be found'
        }

        if ($ExAPropertyBag.ContainsKey($Server) -eq $false)
        {
            Write-Debug "Property Bag has no entries for server '$Server'"
            return $null
        }

        if ($ExAPropertyBag[$Server].ContainsKey($Property) -eq $false)
        {
            Write-Debug "Property bag does not have property '$Property' for '$Server'"
            return $null
        }

        return $ExAPropertyBag[$Server][$Property]
    }
    catch
    {
        Write-Warning "Error $($_.Exception) retrieving '$Property' for '$Server'"
        return $null
    }
}

#this function publishes a value in to the global server property bag
function Set-ExAServerProperty()
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Server,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $Property,
        [parameter(Mandatory=$false)]
        $Value = $null
    )

    Write-Debug "Publishing '$Property' for '$Server'"

    if ($ExAPropertyBag -eq $null)
    {
        throw 'ExAPropertyBag cannot be found'    
    }

    if ($ExAPropertyBag.ContainsKey($Server) -eq $false)
    {
        $ExAPropertyBag.Add($Server, @{})
    }

    $ExAPropertyBag[$Server][$Property] = $Value
}

Function Get-ExARegistryValue()
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $Host,
        [parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryHive]
        $Hive,
        [parameter(Mandatory=$true)]
        [string]
        $Key,
        [parameter(Mandatory=$true)]
        [string]
        $Value,
        [parameter(Mandatory=$false)]
        [object]
        $Default = $null
    )

    ##TODO: Exception Handling
        # Can't connect
        # Access Denied
        # Key doesn't exist
        # Value doesn't exist - default should work for us

    try
    {
        $remoteHive = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Host)

        $hKey = $remoteHive.OpenSubKey($Key, $false)

        $valueData = $hKey.GetValue($Value, $Default)
    }
    catch
    {
        $remoteHive = $null

        $valueData = $null
        
        Write-Error "Unable to get registry value. $($_.Exception.Message)"
    }

    return $valueData
}

#lightweight wrapper over Get-WmiObject with error handling
function Get-ExAWmiObject()
{    
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Class,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Computer,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string[]]        
        $Property,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Filter
    )

    $parameterSet = @{ 
        Computer = $Computer
        Class = $Class
    }

    if ($PSBoundParameters.ContainsKey('Property'))
    {
        $parameterSet.Add('Property', $Property)
    }

    if ($PSBoundParameters.ContainsKey('Filter'))
    {
        $parameterSet.Add('Filter', $Filter)
    }

    ##TODO:
        # Computer down
        # Access Denied

    try
    { 
        Write-Debug "Invoking Get-WmiObject for host '$Computer' and Class '$Class'"
        Get-WmiObject @parameterSet
    }
    catch
    {
        ##TODO
        Write-Warning $_.Exception
        return $null
    }
}

Export-ModuleMember -Function * -Alias * -Variable *