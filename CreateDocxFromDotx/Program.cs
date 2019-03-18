﻿using System;

namespace CreateDocxFromDotx
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var createDocx = new CreateDocx();
                createDocx.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
