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
            



 
            Console.ReadLine();
        }
    }
}
