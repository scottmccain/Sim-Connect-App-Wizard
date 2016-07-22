
function OnFinish(selProj, selObj) {
    try {
        var strProjectPath = wizard.FindSymbol('PROJECT_PATH');
        var strProjectName = wizard.FindSymbol('PROJECT_NAME');

        var strSafeProjectName = CreateSafeName(strProjectName);
        wizard.AddSymbol("SAFE_PROJECT_NAME", strSafeProjectName);
        wizard.AddSymbol("SAFE_APPCLASS_NAME", "C" + strSafeProjectName + "App");
        wizard.AddSymbol("APPLICATION_HEADER", strSafeProjectName + "App.h");
        wizard.AddSymbol('APPLICATION_CLASSFILE', strSafeProjectName + "App.cpp");

        wizard.AddSymbol('MODULE_DEFINITION_FILE', strSafeProjectName + '.def');

        selProj = CreateCustomProject(strProjectName, strProjectPath);
        AddConfig(selProj, strProjectName);
        AddFilters(selProj);

        var InfFile = CreateCustomInfFile();
        AddFilesToCustomProj(selProj, strProjectName, strProjectPath, InfFile);
        PchSettings(selProj);
        InfFile.Delete();

        selProj.Object.Save();
    }
    catch (e) {
        if (e.description.length != 0)
            SetErrorInfo(e);
        return e.number
    }
}

function CreateCustomProject(strProjectName, strProjectPath) {
    try {
        var strProjTemplatePath = wizard.FindSymbol('START_PATH');
        var strProjTemplate = '';
        strProjTemplate = strProjTemplatePath + '\\default.vcxproj';

        var Solution = dte.Solution;
        var strSolutionName = "";
        if (wizard.FindSymbol("CLOSE_SOLUTION")) {
            Solution.Close();
            strSolutionName = wizard.FindSymbol("VS_SOLUTION_NAME");
            if (strSolutionName.length) {
                var strSolutionPath = strProjectPath.substr(0, strProjectPath.length - strProjectName.length);
                Solution.Create(strSolutionPath, strSolutionName);
            }
        }

        var strProjectNameWithExt = '';
        strProjectNameWithExt = strProjectName + '.vcxproj';

        var oTarget = wizard.FindSymbol("TARGET");
        var prj;
        if (wizard.FindSymbol("WIZARD_TYPE") == vsWizardAddSubProject)  // vsWizardAddSubProject
        {
            var prjItem = oTarget.AddFromTemplate(strProjTemplate, strProjectNameWithExt);
            prj = prjItem.SubProject;
        }
        else {
            prj = oTarget.AddFromTemplate(strProjTemplate, strProjectPath, strProjectNameWithExt);
        }
        var fxtarget = wizard.FindSymbol("TARGET_FRAMEWORK_VERSION");
        if (fxtarget != null && fxtarget != "") {
            fxtarget = fxtarget.split('.', 2);
            if (fxtarget.length == 2)
                prj.Object.TargetFrameworkVersion = parseInt(fxtarget[0]) * 0x10000 + parseInt(fxtarget[1])
        }
        return prj;
    }
    catch (e) {
        throw e;
    }
}

function AddFilters(proj) {
    try {
        // Add the folders to your project
        var strSrcFilter = wizard.FindSymbol('SOURCE_FILTER');
        var group = proj.Object.AddFilter('Source Files');
        group.Filter = strSrcFilter;

        var strSrcFilter = wizard.FindSymbol('INCLUDE_FILTER');
        var group = proj.Object.AddFilter('Include Files');
        group.Filter = strSrcFilter;

    }
    catch (e) {
        throw e;
    }
}

function AddConfig(proj, strProjectName) {
    try {
        var config = proj.Object.Configurations("Debug|Win32");
        config.CharacterSet = charSetMBCS;
        config.ConfigurationType = ConfigurationTypes.typeDynamicLibrary;

        var CLTool = config.Tools("VCCLCompilerTool");

        // Warning CLTool.RuntimeLibrary = runtimeLibraryOption.rtSingleThreadedDebug
        // not work with Visual C++ 8 Express
        // Bug corrected with try catch
        try { CLTool.RuntimeLibrary = runtimeLibraryOption.rtMultiThreadedDebug; } catch (ex) { }

        CLTool.Optimization = optimizeOption.optimizeDisabled;
        CLTool.DebugInformationFormat = debugOption.debugEditAndContinue;
        CLTool.WarningLevel = warningLevelOption.warningLevel_3;
        CLTool.AdditionalIncludeDirectories = wizard.FindSymbol('SIMCONNECT_INCLUDE_PATH');

        var strDefines = "WIN32;_WINDOWS;_USRDLL;_DEBUG";

        //strDefines += ";_CONSOLE";

        CLTool.PreprocessorDefinitions = strDefines;

        CLTool.UsePrecompiledHeader = pchOption.pchCreateUsingSpecific;

        var LinkTool = config.Tools("VCLinkerTool");
        LinkTool.ProgramDatabaseFile = "$(outdir)/" + strProjectName + ".pdb";
        LinkTool.GenerateDebugInformation = true;
        LinkTool.LinkIncremental = linkIncrementalYes;
        LinkTool.SubSystem = subSystemWindows;
        LinkTool.OutputFile = "$(outdir)/" + strProjectName + ".dll";
        LinkTool.AdditionalLibraryDirectories = wizard.FindSymbol('SIMCONNECT_LIBRARY_PATH');
        var oldDeps = LinkTool.AdditionalDependencies;
        LinkTool.AdditionalDependencies = oldDeps + " SimConnect.lib";
        LinkTool.ModuleDefinitionFile = wizard.FindSymbol('MODULE_DEFINITION_FILE');

        config = proj.Object.Configurations.Item("Release|Win32");
        config.CharacterSet = charSetMBCS;
        config.ConfigurationType = ConfigurationTypes.typeDynamicLibrary;


        var CLTool = config.Tools("VCCLCompilerTool");

        CLTool.UsePrecompiledHeader = pchOption.pchCreateUsingSpecific;

        CLTool.WarningLevel = warningLevelOption.warningLevel_3;
        CLTool.AdditionalIncludeDirectories = wizard.FindSymbol('SIMCONNECT_INCLUDE_PATH');


        // Warning CLTool.RuntimeLibrary = runtimeLibraryOption.rtSingleThreaded
        // not work with Visual C++ 8 Express
        // Bug corrected with try catch

        try { CLTool.RuntimeLibrary = runtimeLibraryOption.rtMultiThreaded; } catch (ex) { }


        CLTool.InlineFunctionExpansion = expandOnlyInline;

        var strDefines = "WIN32;NDEBUG;_WINDOWS;_USRDLL";

        CLTool.PreprocessorDefinitions = strDefines;

        var LinkTool = config.Tools("VCLinkerTool");
        LinkTool.GenerateDebugInformation = true;
        LinkTool.LinkIncremental = linkIncrementalNo;
        LinkTool.AdditionalLibraryDirectories = wizard.FindSymbol('SIMCONNECT_LIBRARY_PATH');

        LinkTool.SubSystem = subSystemWindows;

        LinkTool.OutputFile = "$(outdir)/" + strProjectName + ".dll";

    }
    catch (e) {
        throw e;
    }
}

function PchSettings(proj) {
    // TODO: specify pch settings
}

function DelFile(fso, strWizTempFile) {
    try {
        if (fso.FileExists(strWizTempFile)) {
            var tmpFile = fso.GetFile(strWizTempFile);
            tmpFile.Delete();
        }
    }
    catch (e) {
        throw e;
    }
}

function CreateCustomInfFile() {
    try {
        var fso, TemplatesFolder, TemplateFiles, strTemplate;
        fso = new ActiveXObject('Scripting.FileSystemObject');

        var TemporaryFolder = 2;
        var tfolder = fso.GetSpecialFolder(TemporaryFolder);
        var strTempFolder = tfolder.Drive + '\\' + tfolder.Name;

        var strWizTempFile = strTempFolder + "\\" + fso.GetTempName();

        var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');
        var strInfFile = strTemplatePath + '\\Templates.inf';
        wizard.RenderTemplate(strInfFile, strWizTempFile);

        var WizTempFile = fso.GetFile(strWizTempFile);
        return WizTempFile;
    }
    catch (e) {
        throw e;
    }
}

function GetTargetName(strName, strProjectName) {
    try {
        // TODO: set the name of the rendered file based on the template filename
        var strTarget = strName;

        if (strName == 'readme.txt')
            strTarget = 'ReadMe.txt';

        if (strName == 'sample.txt')
            strTarget = 'Sample.txt';

        if (strName == 'Module.cpp')
            strTarget = wizard.FindSymbol('SAFE_PROJECT_NAME') + '.cpp';

        if (strName == 'Application.cpp')
            strTarget = wizard.FindSymbol('APPLICATION_CLASSFILE');

        if (strName == 'Application.h')
            strTarget = wizard.FindSymbol('APPLICATION_HEADER');

        if (strName == 'Module.def')
            strTarget = wizard.FindSymbol('MODULE_DEFINITION_FILE');

        return strTarget;
    }
    catch (e) {
        throw e;
    }
}

function AddFilesToCustomProj(proj, strProjectName, strProjectPath, InfFile) {
    try {
        var projItems = proj.ProjectItems

        var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');

        var strTpl = '';
        var strName = '';

        var strTextStream = InfFile.OpenAsTextStream(1, -2);
        while (!strTextStream.AtEndOfStream) {
            strTpl = strTextStream.ReadLine();
            if (strTpl != '') {
                strName = strTpl;
                var strTarget = GetTargetName(strName, strProjectName);
                var strTemplate = strTemplatePath + '\\' + strTpl;
                var strFile = strProjectPath + '\\' + strTarget;

                var bCopyOnly = false;  //"true" will only copy the file from strTemplate to strTarget without rendering/adding to the project
                var strExt = strName.substr(strName.lastIndexOf("."));
                if (strExt == ".bmp" || strExt == ".ico" || strExt == ".gif" || strExt == ".rtf" || strExt == ".css")
                    bCopyOnly = true;
                wizard.RenderTemplate(strTemplate, strFile, bCopyOnly);
                proj.Object.AddFile(strFile);
            }
        }
        strTextStream.Close();
    }
    catch (e) {
        throw e;
    }
}

