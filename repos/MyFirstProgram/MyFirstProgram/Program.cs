using System;
using System.Security.Cryptography;

namespace MyFirstProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            Random random = new Random();
            int num = random.Next(1, 7);

            Console.WriteLine(num);


            Console.ReadKey();
        }
    }
}