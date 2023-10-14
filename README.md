NuGet Feed for CI build: https://ci.appveyor.com/nuget/open-xml-powertools

News
====
This repository was forked from Erik White's as of 01/10/2023. There were over four years of pull requests pending that have been incorporated. There will be a 'pure' release for Net6.0 with all packages and references updated.

This branch has been detached from the original fork, as the intention is to go in a bit of a different direction, which will result in significant changes to the code base, making cross-pull requests difficult.

Moving forward:
1. Remove Stylecop, which isn't being used anyway, and instead gravitate towards best practices SOLID/DRY principles, standard design patterns (see GoF) and using latest C# syntax sugar.
1. Break classes out into separate files that are clean and closed to change.
2. Refactor reusuable code into public or extension methods (this should reduce the code base a lot)
3. Implement local methods rather than countless private methods and arrow methods of doom that make  overaching local of a method indeterminable.
4. Switch from XUnit to MSTests. XUnit has fallen behind in features and a challenge to keep working. NUnit is arguably better, but MSTest is Microsoft, and this is a Microsoft extension framework.
5. Start adding xmldoc and other incode documentation.

Open-XML-PowerTools
===================
The Open XML PowerTools provides guidance and example code for programming with Open XML
Documents (DOCX, XLSX, and PPTX).  It is based on, and extends the functionality
of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

It supports scenarios such as:
- Splitting DOCX/PPTX files into multiple files.
- Combining multiple DOCX/PPTX files into a single file.
- Populating content in template DOCX files with data from XML.
- High-fidelity conversion of DOCX to HTML/CSS.
- High-fidelity conversion of HTML/CSS to DOCX.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Comparing two DOCX files, producing a DOCX with revision tracking markup, and enabling retrieving a list of revisions.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.
- Extracting data (along with formatting) from spreadsheets.

Copyright (c) Microsoft Corporation 2012-2017
Portions Copyright (c) Eric White Inc 2018-2019

Licensed under the MIT License.
See License in the project root for license information.


Open-Xml-PowerTools Content
===========================

Erik White put a lot of work into creating content for Powertools, and it's still the best resources around.

There is a lot of content about Open-Xml-PowerTools at the [Open-Xml-PowerTools Resource Center at OpenXmlDeveloper.org](http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx)

See:
- [DocumentBuilder Resource Center](http://www.ericwhite.com/blog/documentbuilder-developer-center/)
- [PresentationBuilder Resource Center](http://www.ericwhite.com/blog/presentationbuilder-developer-center/)
- [WmlToHtmlConverter Resource Center](http://www.ericwhite.com/blog/wmltohtmlconverter-developer-center/)
- [DocumentAssembler Resource Center](http://www.ericwhite.com/blog/documentassembler-developer-center/)

Build Instructions
==================

**Prerequisites:**

- Visual Studio 2022
- Net 6 SDK, Runtime
- ASP.NET Core Runtime

**Build**
 
 With Visual Studio:

- Open `OpenXmlPowerTools2023.sln` in Visual Studio
- Rebuild the project
- Build the solution.  To validate the build, open the Test Explorer.  Click Run All.
- To run an example, set the example as the startup project, and press F5.

With .NET CLI toolchain:

- Run `dotnet build OpenXmlPowerTools.sln`

Change Log
==========


Version 1.0 : October 14, 2023
- Brought in outstanding valid PRs that introduce many fixes and features. See the initial commit for details.
- Migrated to Net 6.0 and Framework 4.8.
- Significant refactor of project structure to make the project easier to maintain moving forward

Previous Project Fork: https://github.com/OpenXmlDev/Open-Xml-PowerTools
Version 4.6 : November 16, 2020
- Various small bug fixes

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
