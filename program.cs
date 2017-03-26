using System;
using System.IO;

namespace FileApplication
{
   class Program
   {
      static void Main(string[] args)
      {
         // Read and show each line from the file.
         string line = "";
         using (StreamReader sr = new StreamReader("D:\\output.CSV"))
         {
			using (StreamWriter sw = new StreamWriter("E:\\output.CSV"))
			{
				while ((line = sr.ReadLine()) != null)
				{
					sw.WriteLine(line);
				}
			}
         }
      }
   }
}