using System;

using CorelDRAWApplication     = Corel.Interop.VGCore.Application;
using CdrAlignment             = Corel.Interop.VGCore.cdrAlignment;
using CdrTriState              = Corel.Interop.VGCore.cdrTriState;
using CdrFontLine              = Corel.Interop.VGCore.cdrFontLine;
using CdrTextLanguage          = Corel.Interop.VGCore.cdrTextLanguage;
using CdrTextCharSet           = Corel.Interop.VGCore.cdrTextCharSet;
using CdrFileVersion           = Corel.Interop.VGCore.cdrFileVersion;
using CdrUnit                  = Corel.Interop.VGCore.cdrUnit;
using CorelStructCreateOptions = Corel.Interop.VGCore.StructCreateOptions;
using CorelStructSaveAsOptions = Corel.Interop.VGCore.StructSaveAsOptions;
using System.IO;

namespace TestFont
{
  class Program
  {
    static void Main(string[] args)
    {
      string lFontName = args.Length > 0 ? args[0] : "Winter Story";

      CreateText(lFontName);
    }

    static void CreateText( string aFontName )
    {
      Type lPiaType = Type.GetTypeFromProgID("CorelDRAW.Application.22");

      var lApp = Activator.CreateInstance(lPiaType) as CorelDRAWApplication;

      CorelStructCreateOptions lCreateOptions = new CorelStructCreateOptions();

      lCreateOptions.Units = CdrUnit.cdrInch ;
      lCreateOptions.Name  = $"TestFont_{aFontName}";

      var lDoc = lApp.CreateDocumentEx(lCreateOptions);

      Console.WriteLine($"Creating Artistic Text using font: [{aFontName}]");

      var lText = lDoc.ActiveLayer.CreateArtisticTextWide( 0.0, 0.0
                                                         , aFontName
                                                         , CdrTextLanguage.cdrLanguageNone
                                                         , CdrTextCharSet.cdrCharSetMixed
                                                         , aFontName
                                                         , 72
                                                         , CdrTriState.cdrFalse
                                                         , CdrTriState.cdrFalse
                                                         , CdrFontLine.cdrMixedFontLine
                                                         , CdrAlignment.cdrCenterAlignment
                                                         ) ;

      if (  lText.Text.Story.Font != aFontName )
        Console.WriteLine($"ERROR: Text Font changed to [{lText.Text.Story.Font}]");

      string lFileName = $"{Directory.GetCurrentDirectory()}\\TestFont_{aFontName}.cdr";

      Console.WriteLine($"CDR created: [{lFileName}]");

      CorelStructSaveAsOptions lOptions = lApp.CreateStructSaveAsOptions();
      lOptions.Overwrite = true ;
      lOptions.Version = CdrFileVersion.cdrCurrentVersion ;
      lDoc.SaveAs(lFileName, lOptions) ; 
      lDoc.Dirty = false ;
      lDoc.Close();

    }
  }
}
