#Exchange Analyzer

Exchange Analyzer is a PowerShell tool that scans an Exchange Server 2013 or 2016 organization and reports on compliance with best practices.

Exchange Analyzer is currently a beta release seeking feedback and results from real world environments.


###Table of Contents

- [Installing Exchange Analyzer](#installing-exchange-analyzer)
- [Running Exchange Analyzer](#running-exchange-analyzer)
- [Exchange Analyzer Output](#exchange-analyzer-output)
- [Wiki Home Page](https://github.com/cunninghamp/ExchangeAnalyzer/wiki)
	- [Exchange Analyzer Tests](https://github.com/cunninghamp/ExchangeAnalyzer/wiki/Exchange-Analyzer-Tests)
	- [Draft Tests](https://github.com/cunninghamp/ExchangeAnalyzer/wiki/Draft-Tests)
	- [Development Overview](https://github.com/cunninghamp/ExchangeAnalyzer/wiki/Development-Overview)
	- [How to Contribute](https://github.com/cunninghamp/ExchangeAnalyzer/wiki/How-to-Contribute)
- [Credits](#credits)
- [License](#license)

##Installing Exchange Analyzer

Exchange Analyzer consists of a series of PowerShell scripts, PowerShell modules, and XML data files. Installation is performed manually using the following steps.

1. Copy the following files and folders to a computer that has the Exchange 2013/2016 management shell installed. For example, place all of the files and folders in a C:\Scripts\ExchangeAnalyzer folder.

	- Run-ExchangeAnalyzer.ps1
	- Data
	- Modules
	- Tests

2. Copy the folders in the Modules folder to C:\Windows\System32\WindowsPowerShell\v1.0\Modules\

##Running Exchange Analyzer

To run the Exchange Analyzer open an Exchange management shell, navigate to the folder with the script files (e.g. C:\Scripts\ExchangeAnalyzer) and run:

```
.\Run-ExchangeAnalyzer.ps1
```

To see verbose output run:

```
.\Run-ExchangeAnalyzer.ps1 -Verbose
```

##Exchange Analyzer Output

Internet Explorer will automatically open the HTML report when the script has finished. 

[View Sample Report](http://htmlpreview.github.com/?https://github.com/cunninghamp/ExchangeAnalyzer/blob/master/SampleReport.html) (when repo is public)

##Credits

- Paul Cunningham ([Blog](http://exchangeserverpro.com) | [Twitter](https://twitter.com/exchservpro))
- Mike Crowley
- Michael B Smith
- Brian Desmond
- Damian Scoles

##License

Exchange Analyzer is released under the MIT license (MIT). Full text of the license is available [here](https://github.com/cunninghamp/ExchangeAnalyzer/blob/master/LICENSE).