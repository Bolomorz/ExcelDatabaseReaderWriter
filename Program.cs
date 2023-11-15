// See https://aka.ms/new-console-template for more information
using ConsoleApp1;
using Mysqlx.Resultset;

//-
//--
//---
//Database
DatabaseReaderWriter fahrradverleihdb = new DatabaseReaderWriter("server=localhost; database=fahrradverleihdb; user = root");
string update = "UPDATE tbl_fahrrad SET FrdMarke = 'Giant' WHERE FrdNr = 1";
string? com = fahrradverleihdb.CommandNonQuery(update);
if(com is null)
{
    Console.WriteLine(update + " was successful!");
}
else
{
    Console.WriteLine(com);
}

string select = "SELECT f.FrdNr, f.FrdMarke FROM tbl_fahrrad f";
var query = fahrradverleihdb.CommandQuery(select);
if(query.errormessage is null && query.rows is not null)
{
    Console.WriteLine(select + " was successful!");
    foreach (var row in query.rows)
    {
        Console.WriteLine("FrdNr: " + row[0] + "; FrdMarke: " + row[1]);
    }
}
else
{
    Console.WriteLine(query.errormessage);
}
//---
//--
//-

//-
//--
//---
//Excel
ExcelReaderWriter example = new ExcelReaderWriter(@"C:\Users\domin\OneDrive\Dokumente\testexcel.xlsx");
string? message = example.WriteCell(1,1, "example");
if(message is null)
{
    Console.WriteLine("WriteCell was successful!");
}
else
{
    Console.WriteLine(message);
}

Tuple<string?, object?> ret2 = example.ReadCell(1,1);
if(ret2.Item1 is null && ret2.Item2 is not null)
{
    Console.WriteLine("ReadCell was successful!");
    Console.WriteLine(ret2.Item2.ToString());
}
else
{
    Console.WriteLine(ret2.Item1);
}
//---
//--
//-
