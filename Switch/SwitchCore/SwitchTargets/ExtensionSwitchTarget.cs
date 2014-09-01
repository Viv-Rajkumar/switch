using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.VCCodeModel;
using Microsoft.VisualStudio.VCProjectEngine;

namespace SwitchCore.SwitchTargets {
  /// <summary>
  ///   An Extension DoSwitch Target switches files based on extension.
  /// </summary>
  public class ExtensionSwitchTarget : ISwitchTarget {
    /// <summary>
    ///   Initializes a new instance of the <see cref="ExtensionSwitchTarget" /> class.
    /// </summary>
    /// <param name="from">From.</param>
    /// <param name="to">To.</param>
    public ExtensionSwitchTarget(string from, string to) {
      From = from;
      To = to;
    }

    /// <summary>
    ///   Gets from.
    /// </summary>
    public string From { get; private set; }

    /// <summary>
    ///   Gets to.
    /// </summary>
    public string To { get; private set; }

    /// <summary>
    ///   Switches the specified document.
    /// </summary>
    /// <param name="application">The application.</param>
    /// <param name="activeDocument">The active document.</param>
    /// <returns>
    ///   True if switched successfully.
    /// </returns>
    public bool DoSwitch(DTE2 application, Document activeDocument) {
      var folderPath = Path.GetDirectoryName(activeDocument.FullName);
      var fileName = Path.GetFileNameWithoutExtension(activeDocument.FullName);
      var extension = Path.GetExtension(activeDocument.FullName);

      Debug.Assert(folderPath != null, "folderPath != null");
      Debug.Assert(fileName != null, "fileName != null");

      if (extension == null || extension.Replace(".", "") != From) {
        return false;
      }

      var mappedPath = Path.Combine(folderPath, fileName) + "." + To;
      if (SwitchHelper.TryOpenDocument(application, mappedPath)) {
        return true;
      }

      try {
        var currentProject = activeDocument.ProjectItem.ContainingProject;
        var currentFile = activeDocument.ProjectItem.Object as VCFile;

        if (currentFile.FileType == eFileType.eFileTypeCppCode) {
          var configName = currentProject.ConfigurationManager.ActiveConfiguration.ConfigurationName;
          mappedPath = ResolveHeaderPath(activeDocument, fileName, currentFile, configName);
        } else {
          mappedPath = ResolveSourcePath(activeDocument, fileName);
        }

        return SwitchHelper.TryOpenDocumentFromProjectIncludes(application, mappedPath);
      } catch (Exception) {
        return false;
      }
    }

    /// <summary>
    ///   Determines whether this instance can switch given the specified document.
    /// </summary>
    /// <param name="application">The application.</param>
    /// <param name="activeDocument">The active document.</param>
    /// <returns>
    ///   <c>true</c> if this instance can switch the specified application; otherwise, <c>false</c>.
    /// </returns>
    public bool CanSwitch(DTE2 application, Document activeDocument) {
      //  For now we always enable the command.
      return true;
    }

    private string ResolveHeaderPath(Document activeDocument, string fileName, VCFile currentFile, string configName) {
      var codeModel = activeDocument.ProjectItem.FileCodeModel as VCFileCodeModel;
      CodeElement codeElement =
        codeModel.Includes.Cast<CodeElement>().
          FirstOrDefault(element => element.FullName.EndsWith(fileName + "." + To));
      var fileConfig = currentFile.FileConfigurations.Item(configName) as VCFileConfiguration;
      var fileRule = fileConfig.Tool as IVCRulePropertyStorage;
      var includeDirectories = fileRule.GetEvaluatedPropertyValue("AdditionalIncludeDirectories").Split(';');
      foreach (
        var headerPath in
          includeDirectories.Select(
            includeDirectory =>
              Path.Combine(includeDirectory, codeElement.FullName).
                Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)).
            Where(File.Exists)) {
        return headerPath;
      }

      throw new FileNotFoundException();
    }

    private string ResolveSourcePath(Document activeDocument, string fileName) {
      var currentVCProject = activeDocument.ProjectItem.ContainingProject.Object as VCProject;
      foreach (VCFile file in currentVCProject.Files) {
        if (file.FullPath.EndsWith(fileName + "." + To)) {
          return file.FullPath;
        }
      }
      throw new FileNotFoundException();
    }
  }
}
