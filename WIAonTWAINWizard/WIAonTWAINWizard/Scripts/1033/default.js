
function OnFinish(selProj, selObj)
{
  try 
	{
	  CreateDBGFile();
		var strProjectPath = wizard.FindSymbol('PROJECT_PATH');
		var strProjectName = wizard.FindSymbol('PROJECT_NAME');
		var strVID = wizard.FindSymbol('VID');
		var strPID = wizard.FindSymbol('PID');
		var strSCSI_VID = wizard.FindSymbol('SCSI_VID');
		var strSCSI_PID = wizard.FindSymbol('SCSI_PID');
		var strWIAdrvName = String(wizard.FindSymbol('WIA_DRIVER'));
    strWIAdrvName = strWIAdrvName.toLowerCase();
    var lastLocation = strWIAdrvName.indexOf(".dll")
    if ( lastLocation >= 0 )
    {
        strWIAdrvName = strWIAdrvName.substr( 0, lastLocation );
    }
    strWIAdrvName += "_UI.dll";
    wizard.AddSymbol("WIA_DRIVER_UI", strWIAdrvName);

    if(strVID!="" && strPID!="")
    {
		  wizard.AddSymbol("USB",1);
    }
    else
    {
		  wizard.AddSymbol("USB",0);
    }

    if(strSCSI_VID!="" && strSCSI_PID!="")
    {
		  wizard.AddSymbol("SCSI",1);
    }
    else
    {
		  wizard.AddSymbol("SCSI",0);
    }

		var strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS1",
           wizard.FormatGuid(strAppGuid, 0));
		wizard.AddSymbol("GUID_CLASS2",
           wizard.FormatGuid(strAppGuid, 1));

		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_UI1",
           wizard.FormatGuid(strAppGuid, 0));
		wizard.AddSymbol("GUID_CLASS_UI2",
           wizard.FormatGuid(strAppGuid, 2));

		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_PR86",
           wizard.FormatGuid(strAppGuid, 0));
		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_PA86",
           wizard.FormatGuid(strAppGuid, 0));
		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_UP86",
           wizard.FormatGuid(strAppGuid, 0));
		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_PR64",
           wizard.FormatGuid(strAppGuid, 0));
		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_PA64",
           wizard.FormatGuid(strAppGuid, 0));
		strAppGuid = wizard.CreateGuid();
		wizard.AddSymbol("GUID_CLASS_UP64",
           wizard.FormatGuid(strAppGuid, 0));

		var today = new Date();
		var strDate="";
		if (today.getMonth() < 9) 
    {
      strDate = "0";
    }
    strDate += (today.getMonth()+1).toString();
    strDate += "/";
    if (today.getDate() < 10) 
    {
      strDate += "0";
    }
    strDate += today.getDate().toString();
    strDate += "/" + today.getFullYear().toString();
    wizard.AddSymbol("DATE", strDate);
    //		09/20/2004
    wizard.AddSymbol("YEAR", today.getFullYear().toString());

    var Solution = dte.Solution;
    var strSolutionName = "";
    Solution.Close();
    strSolutionName = wizard.FindSymbol("VS_SOLUTION_NAME");
    if (strSolutionName.length) 
    {
      var strSolutionPath = strProjectPath.substr(0, strProjectPath.length - strProjectName.length);
      Solution.Create(strSolutionPath, strSolutionName);
    }

		selProj = CreateWIAdriverProject(strProjectName, strProjectPath);
		var InfFile = CreateCustomInfFile("WIAdriverTemplates.inf");
		var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');
		AddFilesToCustomProj(selProj,strTemplatePath, strProjectPath, InfFile, true);
		InfFile.Delete();
		selProj.Object.Save();

    strProjectPath = strSolutionPath + "WIA_UI";
    strTemplatePath = strTemplatePath + "\\WIA_UI";
		selProj = CreateWIA_UIProject(strProjectPath);
		InfFile = CreateCustomInfFile("WIAdriverUITemplates.inf");
		AddFilesToCustomProj(selProj, strTemplatePath, strProjectPath, InfFile, false);
		InfFile.Delete();

		strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');
    strProjectPath = strSolutionPath + "Installer";
    strTemplatePath = strTemplatePath + "\\Installer";
		selProj = CreateMergeModuleProjectProject(strProjectPath);
		InfFile = CreateCustomInfFile("MergeModuleTemplates.inf");
		AddFilesToCustomProj(selProj, strTemplatePath, strProjectPath, InfFile, false);
		InfFile.Delete();

		RenderINFFiles(strProjectPath);
		RenderDriverFiles(strProjectPath);
		selProj = CreateInstallerProject(strSolutionPath + "Installer\\");
		selProj = Create64bitInstallerProject(strSolutionPath + "Installer\\");
		
	}
	catch(e)
	{
		if (e.description.length != 0)
			SetErrorInfo(e);
		return e.number
	}
}
function DelFile(strFileName) 
{
  try 
  {
    var fso;
    fso = new ActiveXObject('Scripting.FileSystemObject');

    if (fso.FileExists(strFileName)) 
    {
      var tmpFile = fso.GetFile(strFileName);
      tmpFile.Delete();
    }
  }
  catch (e) 
  {
    throw e;
  }
}


function CreateWIAdriverProject(strProjectName, strProjectPath)
{
	try
	{
		var strProjTemplatePath = wizard.FindSymbol('PROJECT_TEMPLATE_PATH');
		var strProjTemplate = '';
		strProjTemplate = strProjTemplatePath + '\\wiadriver.vcproj';
		
		var strDestPath = strProjectPath + "\\temp";
	  var strWizTempFile = strDestPath + "\\temp.vcproj";
	  wizard.RenderTemplate(strProjTemplate, strWizTempFile, false);

		var InfFile = CreateCustomInfFile("WIAdriverTemplatesCopy.inf");
		var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');

		AddFilesToCustomProj(0,strTemplatePath, strDestPath, InfFile, false);
		InfFile.Delete();

		var strProjectNameWithExt = '';
		strProjectNameWithExt = strProjectName + '.vcproj';

		var oTarget = wizard.FindSymbol("TARGET");
		var prj;

		if (wizard.FindSymbol("WIZARD_TYPE") == vsWizardAddSubProject)  // vsWizardAddSubProject
		{
		  var prjItem = oTarget.AddFromTemplate(strWizTempFile, strProjectNameWithExt);
			prj = prjItem.SubProject;
		}
		else
		{
		  prj = oTarget.AddFromTemplate(strWizTempFile, strProjectPath, strProjectNameWithExt);
		}
   var fso;
   fso = new ActiveXObject('Scripting.FileSystemObject');
   fso.DeleteFolder(strDestPath);
		return prj;
	}
	catch(e)
	{
		throw e;
	}
}
function CreateWIA_UIProject(strProjectPath)
{
	try
	{
		var strProjTemplatePath = wizard.FindSymbol('PROJECT_TEMPLATE_PATH');
		var strProjTemplate = '';
		strProjTemplate = strProjTemplatePath + '\\WIA_UI\\WIA_UI.vcproj';

		var strWizTempFile = strProjectPath + "\\WIA_UI.vcproj";
	  wizard.RenderTemplate(strProjTemplate, strWizTempFile, false);

	  var prj = dte.Solution.AddFromFile(strWizTempFile);
		
		return prj;

	}
	catch(e)
	{
		throw e;
	}}
function CreateMergeModuleProjectProject(strProjectPath)
{
	try
	{
		var strProjTemplatePath = wizard.FindSymbol('PROJECT_TEMPLATE_PATH');
		var strProjTemplate = '';
		strProjTemplate = strProjTemplatePath + '\\WIADriverMSM.vdproj';

		var strWizTempFile = strProjectPath + "\\WIADriverMSM.vdproj";
	  wizard.RenderTemplate(strProjTemplate, strWizTempFile, false);

	  var prj = dte.Solution.AddFromFile(strWizTempFile);
		
		return prj;
	}
	catch(e)
	{
		throw e;
	}}
function CreateInstallerProject(strProjectPath)
{
	try
	{
		var strProjTemplatePath = wizard.FindSymbol('PROJECT_TEMPLATE_PATH');
		var strProjTemplate = '';
		strProjTemplate = strProjTemplatePath + '\\Installer.vdproj';

		var strWizTempFile = strProjectPath + "\\Installer.vdproj";
	  wizard.RenderTemplate(strProjTemplate, strWizTempFile, false);

	  var prj = dte.Solution.AddFromFile(strWizTempFile);
		
		return prj;
	}
	catch(e)
	{
		throw e;
	}}

function Create64bitInstallerProject(strProjectPath) {
  try {
    var strProjTemplatePath = wizard.FindSymbol('PROJECT_TEMPLATE_PATH');
    var strProjTemplate = '';
    strProjTemplate = strProjTemplatePath + '\\Installer64.vdproj';

    var strWizTempFile = strProjectPath + "\\Installer64.vdproj";
    wizard.RenderTemplate(strProjTemplate, strWizTempFile, false);

    var prj = dte.Solution.AddFromFile(strWizTempFile);

    return prj;
  }
  catch (e) {
    throw e;
  }
}
function CreateCustomInfFile(strInfFileName)
{
	try
	{
		var fso, TemplatesFolder, TemplateFiles, strTemplate;
		fso = new ActiveXObject('Scripting.FileSystemObject');

		var TemporaryFolder = 2;
		var tfolder = fso.GetSpecialFolder(TemporaryFolder);
		var strTempFolder = tfolder.Drive + '\\' + tfolder.Name;

		var strWizTempFile = strTempFolder + "\\" + fso.GetTempName();

		var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');
		var strInfFile = strTemplatePath +"\\"+ strInfFileName;
		wizard.RenderTemplate(strInfFile, strWizTempFile);
		var WizTempFile = fso.GetFile(strWizTempFile);
		return WizTempFile;
	}
	catch(e)
	{
		throw e;
	}
}

function AddFilesToCustomProj(proj, strTemplatePath, strProjectPath, InfFile, bAdd)
{
	try
	{
		var strTpl = '';
		var strName = '';

		var strTextStream = InfFile.OpenAsTextStream(1, -2);
		while (!strTextStream.AtEndOfStream)
		{
			strTpl = strTextStream.ReadLine();
			if (strTpl != '')
			{
				strName = strTpl;
				var strTemplate = strTemplatePath + '\\' + strTpl;
				var strFile = strProjectPath + '\\' + strName;

				var bCopyOnly = false;  //"true" will only copy the file from strTemplate to strTarget without rendering/adding to the project
				var strExt = strName.substr(strName.lastIndexOf("."));
        strExt = strExt.toLowerCase();
				if (strExt == ".bmp" || strExt == ".ico" || strExt == ".gif" || strExt == ".rtf" || strExt == ".css" || strExt == ".lib"|| strExt == ".exe" || strExt == ".dll"  || strExt == ".dll64" || strExt == ".msm"  || strExt == ".mst")
					bCopyOnly = true;
				wizard.RenderTemplate(strTemplate, strFile, bCopyOnly);
        if(bAdd)
        {
				  proj.Object.AddFile(strFile);
        }
			}
		}
		strTextStream.Close();
	}
	catch(e)
	{
		throw e;
	}
}

function RenderINFFiles(strProjectPath)
{
	try
	{
		var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');

		var strName = "DriverFiles\\wiadriver.inf"
		var strTemplate = strTemplatePath + '\\Installer\\' + strName;
		var strFile = strProjectPath + "\\DriverFiles\\" + wizard.FindSymbol('INF_NAME');
  	wizard.RenderTemplate(strTemplate, strFile, false);
	}
	catch(e)
	{
		throw e;
	}
}

function RenderDriverFiles(strProjectPath)
{
	try
	{
		var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');

		var strName = "DriverFiles\\YourCATfile.cat"
		var strTemplate = strTemplatePath + '\\Installer\\' + strName;
		var strFile = strProjectPath + "\\DriverFiles\\" + wizard.FindSymbol('CAT_NAME');
  	wizard.RenderTemplate(strTemplate, strFile, false);

		strName = "DriverFiles\\YourWIAdriver.dll"
		strTemplate = strTemplatePath + '\\Installer\\' + strName;
		strFile = strProjectPath + "\\DriverFiles\\" + wizard.FindSymbol('WIA_DRIVER');
  	wizard.RenderTemplate(strTemplate, strFile, false);
		strName = "DriverFiles\\YourWIAdriver.dll64"
		strTemplate = strTemplatePath + '\\Installer\\' + strName;
		strFile = strProjectPath + "\\DriverFiles\\" + wizard.FindSymbol('WIA_DRIVER')+"64";
  	wizard.RenderTemplate(strTemplate, strFile, false);

		strName = "DriverFiles\\YourWIAdriver_UI.dll"
		strTemplate = strTemplatePath + '\\Installer\\' + strName;
		strFile = strProjectPath + "\\DriverFiles\\" + wizard.FindSymbol('WIA_DRIVER_UI');
  	wizard.RenderTemplate(strTemplate, strFile, false);
		strName = "DriverFiles\\YourWIAdriver_UI.dll64"
		strTemplate = strTemplatePath + '\\Installer\\' + strName;
		strFile = strProjectPath + "\\DriverFiles\\" + wizard.FindSymbol('WIA_DRIVER_UI')+"64";
  	wizard.RenderTemplate(strTemplate, strFile, false);
	}
	catch(e)
	{
		throw e;
	}
}
// ------------------------------  DEBUG HELPERS -----------------------------


var debugON = true;

// don't panic.. the script will check for the existence of the folder below
// before attempting to use it
var debugFolder = "C:\\WizDebug";

var debugFile = "WizLog.txt";

var debugInitialized = false;

/**
*   This function attempts to create a log file for debugging purposes. Because
*   we are JavaScript, we are running in a sandbox and don't have a lot of choices
*   about where we can create files.
*/
function CreateDBGFile() {

  debugInitialized = true;

  if (debugON != true) return;

  var oFSO;
  oFSO = new ActiveXObject("Scripting.FileSystemObject");

  if (oFSO.DriveExists(debugFolder == false)) {
    debugON = false;
    return;
  }

  if (oFSO.FolderExists(debugFolder) == false) {
    debugON = false;
    return;
  }

  debugFile = debugFolder + "\\" + debugFile;

  var oStream = oFSO.CreateTextFile(debugFile, true, true);
  oStream.Close();
}

/**
* This function is used to output debug info to a file so we can debug the script.
*/
function DebugOut(str) {
  if (debugInitialized == false) {
    CreateDBGFile();
  }

  if (debugON != true) return;

  var oFSO;
  oFSO = new ActiveXObject("Scripting.FileSystemObject");

  var oStream = oFSO.OpenTextFile(debugFile, 8, true, -1);
  oStream.Write(str + "\r\n");
  oStream.Close();
  return;
}
