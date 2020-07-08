# excel-plugin-testing

This solution exists to test free plugins for creating word and excel files.

For building Excel Files, ClosedXml is being used which wraps the DocumentFormat.OpenXml library and simplifies the creating of excel files.

For building Word Files, DocumentFormat.OpenXml is being used directly.

This solution has to work in an Azure WebApp where there are no access to office applications.  This has been deployed via GitHub and Kudu to the following location which is IP locked: https://stotty-excel-creation-test.azurewebsites.net/