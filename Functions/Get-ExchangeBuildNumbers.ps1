Function Get-ExchangeBuildNumbers()
{

    $URL = "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.160).aspx"

    $WebPage = Invoke-WebRequest -Uri $URL

    #Reference: Lee Holmes article on extracting tables from web pages
    #Link: #http://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/

    $tables = @($WebPage.Parsedhtml.getElementsByTagName("TABLE"))

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

    return $ExchangeBuildNumbers

}

$buildnumbers = Get-ExchangeBuildNumbers

$buildnumbers

#Test some build number matching
#$build1 = "15.01.0225.042"
#$build2 = "15.00.1044.025"
#$productnames = @($buildnumbers | Where {($_."Product name" -match "Exchange") -and ($_."Build number").Substring(0,4) -eq $build1.Substring(0,4)})
#$productnames