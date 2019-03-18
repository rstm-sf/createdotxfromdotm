using System;

namespace CreateDocxFromDotx
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var createDocx = new CreateDocx();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                Console.WriteLine("\nPress Enter to continue…");
                Console.ReadLine();
            }
        }
    }
}
