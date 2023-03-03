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
    Console.WriteLine(update + " was successfull!");
}
else
{
    Console.WriteLine(com);
}

string select = "SELECT f.FrdNr, f.FrdMarke FROM tbl_fahrrad f";
Tuple<string, List<List<string>>> ret = fahrradverleihdb.Select(select, 2);
if(ret.Item1 == string.Empty)
{
    Console.WriteLine(select + " was successfull!");
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
ExcelReaderWriter example = new ExcelReaderWriter(@"U:\MyExcel.xlsx");
example.WriteCell("A1", "example");
string read = example.ReadCell("A1");
Console.WriteLine(read);
example.SaveAndDispose();
//---
//--
//-
