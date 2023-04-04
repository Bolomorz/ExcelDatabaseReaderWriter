// See https://aka.ms/new-console-template for more information
using ConsoleApp1;
using Mysqlx.Resultset;

//-
//--
//---
//Database
DatabaseReaderWriter fahrradverleihdb = new DatabaseReaderWriter("localhost", "fahrradverleihdb", "root");
string update = "UPDATE tbl_fahrrad SET FrdMarke = 'Giant' WHERE FrdNr = 1";
string com = fahrradverleihdb.Command(update);
if(com == string.Empty)
{
    Console.WriteLine(update + " was successful!");
}
else
{
    Console.WriteLine(com);
}

string select = "SELECT f.FrdNr, f.FrdMarke FROM tbl_fahrrad f";
Tuple<string, List<List<string>>> ret = fahrradverleihdb.Select(select, 2);
if(ret.Item1 == string.Empty)
{
    Console.WriteLine(select + " was successful!");
    foreach (List<string> row in ret.Item2)
    {
        Console.WriteLine("FrdNr: " + row[0] + "; FrdMarke: " + row[1]);
    }
}
else
{
    Console.WriteLine(ret.Item1);
}
//---
//--
//-

//-
//--
//---
//Excel
ExcelReaderWriter example = new ExcelReaderWriter(@"C:\Users\domin\OneDrive\Dokumente\testexcel.xlsx");
string message = example.WriteCell(1,1, "example");
if(message == string.Empty)
{
    Console.WriteLine("WriteCell was successful!");
}
else
{
    Console.WriteLine(message);
}

Tuple<string, object> ret2 = example.ReadCell(1,1);
if(ret2.Item1 == string.Empty)
{
    Console.WriteLine("ReadCell was successful!");
    Console.WriteLine(ret2.Item2);
}
else
{
    Console.WriteLine(ret2.Item1);
}
//---
//--
//-
