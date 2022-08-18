# DataImporterServiceBroker
Nintex K2 4.7 / Five Service Broker to import Excel / CSV files

This is a service broker which allows you to upload table data from an Excel or CSV file to a SmartObject.  

It uses the OpenXML SDK for Office when doing the Excel importing and utilises the samples from [MSDN](http://msdn.microsoft.com/en-us/library/office/gg575571(v=office.15).aspx).

The [GenericParser](https://github.com/AndrewRissing/GenericParsing) component was used for CSV importing.

This was tested on K2 versions 4.7 and 5.4. The solution was built using Visual Studio 2015.

Note that this service broker was created as a proof of concept and is provided as-is.  It has been tested to verify functionality but has not gone through a full QA cycle or tested under heavy load.

It is based upon the Excel Service broker https://community.nintex.com/t5/K2-Five-blackpearl/Excel-Import-Service-Broker/ta-p/177090 and has added a number of new features added based upon the comments in the discussion thread.

New arcade page: https://community.nintex.com/t5/K2-Five-blackpearl/Excel-CSV-Import-Service-Broker/ta-p/176973
