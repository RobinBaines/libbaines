# libbaines
The framework is described in detail in Utilities.pdf.
Here is a summary:
The Utilities and Excel interface libraries provide functionality for .Net applications:
It includes 
  Filtering, Sorting and Editing of data,
  Segregation of Duty,
  Column Design,
  Multi-language support,
  User Logging,
  User Blocking after periods of inactivity,
  Timer initiated termination of an Application,
  Semaphore support for fast updating.
  Flexible connection support for switching between Live and Test databases.
  Filtered and sorted data export to Excel.
  Help textbox in every form.

To create a MSQL database see SQL\readme.txt.
TestApp/TestApp.sln is an example application using the framework.

Modification 20241109 
Updated Utilities, ExcelInterface and TestApp to Framework 4.8.
Added ClosedXML as an alternative to Excel automation in ExcelInterface dll.
TestApp has been modified:
  Use ClosedXML instead of Excel automation.
  Added a form called frmUsrLog2 to TestApp. The vb file includes a detailed description of the steps to create a new form, which inherits from frmStandard and how to include a Utilities.dgvEnter DataGridView.
  Improved TestAppMain.vb.

The GenUi project is now public on this Git site. This can be used to generate dgvEnter code used in TestApp.frmUsrLog2. 


