# DataImporterServiceBroker
K2 Service Broker to import Excel / CSV files

This is a service broker which allows you to upload table data from an Excel or CSV file to a SmartObject.  

It uses the OpenXML SDK for Office when doing the Excel importing and utilises the samples from [MSDN](http://msdn.microsoft.com/en-us/library/office/gg575571(v=office.15).aspx).

The [GenericParser](https://github.com/AndrewRissing/GenericParsing) component was used for CSV importing.

This was tested on K2 versions 4.7 and 5.4. The solution was built using Visual Studio 2015.

Note that this service broker was created as a proof of concept and is provided as-is.  It has been tested to verify functionality but has not gone through a full QA cycle or tested under heavy load.

It is based upon the Excel Service broker https://community.k2.com/t5/K2-blackpearl/Excel-Import-Service-Broker/ba-p/65814 and has added a number of new features added based upon the comments in the discussion thread.

New discussion thread: https://community.k2.com/t5/K2-blackpearl/Excel-CSV-Import-Service-Broker/ba-p/117170

By using this component you are agreeing to the K2 Download License: https://community.k2.com/html/assets/SourceCodeTechnologyHoldingsInc.DownloadLicense.pdf
