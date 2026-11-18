using System;
using CalcRapido;

public static class TestMain
{
    public static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        return SelfTest.Run();
    }
}
