using System;
using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;

namespace SwitchCore {
  /// <summary>
  ///   Helper functions for DoSwitch functionality.
  /// </summary>
  public static class SwitchHelper {
    /// <summary>
    ///   Tries the open document.
    /// </summary>
    /// <param name="application">The application.</param>
    /// <param name="path">The path.</param>
    /// <returns></returns>
    public static bool TryOpenDocument(DTE2 application, string path) {
      //  Go through each document in the solution.
      foreach (Document document in application.Documents) {
        if (String.CompareOrdinal(document.FullName, path) == 0) {
          //  The document is open, we just need to activate it.
          if (document.Windows.Count > 0) {
            document.Activate();
            return true;
          }
        }
      }

      //  The document isn't open - does it exist?
      if (File.Exists(path)) {
        try {
          application.Documents.Open(path, "Text", false);
          return true;
        } catch (Exception) {
          //  We can't open the doc.
          return false;
        }
      }

      //  We couldn't open the document.
      return false;
    }

    public static bool TryOpenDocumentFromProjectIncludes(DTE2 application, string path) {
      foreach (var document in from Document document in application.Documents
        where String.CompareOrdinal(document.FullName, path) == 0
        where document.Windows.Count > 0
        select document) {
        document.Activate();
        return true;
      }

      if (!File.Exists(path)) {
        return false;
      }

      try {
        application.Documents.Open(path, "Text", false);
        return true;
      } catch (Exception) {
        return false;
      }
    }
  }
}
