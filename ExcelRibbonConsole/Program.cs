using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRibbonAddin
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel App");
            Console.WriteLine(" *If you are not prompted to Enable Macros and nothing else happens, your security level is probably on High. Set it to Medium.");
            Console.WriteLine(@" *If you get a message indicating that the.xll file is not recognized(and even opening it as text), you might not have the.NET Framework 2.0 installed.Install it. 
                                 Otherwise Excel is using .Net version 1.1 by default.To change this, refer back to the prerequisites section.");

            Console.WriteLine(@"*If Excel crashes with an unhandled exception, an access violation or some other horrible error, either during loading or when running the function, please let me know.
                                This shouldn't happen, and I would like to know if it does.");
            Console.WriteLine(@" * If a form appears with the title 'ExcelDna Compilation Errors' then there were some errors trying to compile the code in the.dna file. Check that you have put the right code into the .dna file.");
            Console.WriteLine(@" *If Excel prompts for Enabling Macros, and then the function does not work and does not appear in the function wizard, you might not have the right filename for the.dna file. The prefix should be the same as the.xll file and it should be in the same directory.");


            var Lookup = @"            id2_Item = Reader.NameTable.Add(@'');
                                        id14_CustomUI = Reader.NameTable.Add('CustomUI');
                                        id24_LoadFromBytes = Reader.NameTable.Add('LoadFromBytes');
                                        id16_Path = Reader.NameTable.Add('Path');
                                        id17_Pack = Reader.NameTable.Add('Pack');
                                        id7_Language = Reader.NameTable.Add('Language');
                                        id20_ExplicitRegistration = Reader.NameTable.Add('ExplicitRegistration');
                                        id11_ExternalLibrary = Reader.NameTable.Add('ExternalLibrary');
                                        id3_Name = Reader.NameTable.Add('Name');
                                        id6_CreateSandboxedAppDomain = Reader.NameTable.Add('CreateSandboxedAppDomain');
                                        id10_DefaultImports = Reader.NameTable.Add('DefaultImports');
                                        id9_DefaultReferences = Reader.NameTable.Add('DefaultReferences');
                                        id12_Project = Reader.NameTable.Add('Project');
                                        id18_AssemblyPath = Reader.NameTable.Add('AssemblyPath');
                                        id22_SourceItem = Reader.NameTable.Add('SourceItem');
                                        id13_Reference = Reader.NameTable.Add('Reference');
                                        id23_TypeLibPath = Reader.NameTable.Add('TypeLibPath');
                                        id5_ShadowCopyFiles = Reader.NameTable.Add('ShadowCopyFiles');
                                        id15_Image = Reader.NameTable.Add('Image');
                                        id19_ExplicitExports = Reader.NameTable.Add('ExplicitExports');
                                        id1_DnaLibrary = Reader.NameTable.Add('DnaLibrary');
                                        id4_RuntimeVersion = Reader.NameTable.Add('RuntimeVersion');
                                        id8_CompilerVersion = Reader.NameTable.Add('CompilerVersion');
                                        id21_ComServer = Reader.NameTable.Add('ComServer');";
          Console.WriteLine(string.Format(" xml is looking for these keywords so they have to match and be of the same case {1}", Lookup));
 
            Console.ReadLine();
        }
    }
}
