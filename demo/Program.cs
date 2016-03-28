using System;
using System.Collections.Generic;

using ObjectPrinterLib;

namespace demo
{
	class Program
	{
		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			
			// TODO: Implement Functionality Here
			
			var people =new List<Person> ();
			people.Add (new Person () { Age = 88, Name = "Benjamin Graham" } );
			people.Add (new Person () { Age = 12, Name = "Jean Claude VanDamme" } );
			
			Console.WriteLine ("Exporting People to Excel");
			GenericListExport.ExportExcel<Person> ( people, "Name;Age", "myPeople.xlsx");
		
			Console.WriteLine ("Exporting People to Html");
			Console.WriteLine ( GenericListExport.ExportHtml<Person> ( people, "Name;Age" ));
		
			
		
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}