# ExchangeAnalyzer

##Installing

1. Copy the following files and folders to a computer that has the Exchange 2013/2016 management shell installed. For example, place all of the files and folders in a C:\Scripts\ExchangeAnalyzer folder.

	- Run-ExchangeAnalyzer.ps1
	- Data
	- Modules
	- Tests

2. Copy the folders in the Modules folder to C:\Windows\System32\WindowsPowerShell\v1.0\Modules\

##Running

To run the Exchange Analyzer open an Exchange management shell, navigate to the folder with the script files (e.g. C:\Scripts\ExchangeAnalyzer) and run:

```
.\Run-ExchangeAnalyzer.ps1
```

To see verbose output run:

```
.\Run-ExchangeAnalyzer.ps1 -Verbose
```

##Output

Internet Explorer will automatically open the HTML report when the script has finished. 
 

[View Sample Report](http://htmlpreview.github.com/?https://github.com/cunninghamp/ExchangeAnalyzer/blob/master/SampleReport.html) (when repo is public)
