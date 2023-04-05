using MongoExcelMigration.Modelos;
using NPOI.SS.Formula.Functions;

//Versión: 1.0
//Fecha: 04 de Abril de 2023
//Autor: ASLOGIC S.A. DE C.V.
//Desarrollador: Ammi Jatziry Wang Almazán
//Módulo: Program.
//Descripción: Interfaz para el usuario, para agregar a el excel a MongoDB.
//Historial de cambios:
//04 de Abril de 2023: .

bool flag = true;

while (flag)
{
    Console.WriteLine("Elige una opcion a realizar:\n1.-Migrar Excel a Mongo\n2.-Salir");
    try
    {
        int option = int.Parse(Console.ReadLine());
        switch (option)
        {
            case 1:
                {
                    try
                    {
                        Console.WriteLine("Ingrese el nombre del archivo de Excel (Ejemplo.xlsx)");
                        String path = Console.ReadLine();
                        Console.WriteLine("Ingrese el valor de la fila en que comienzan los encabezados de la tabla (Ejem. 0)");
                        int initTable = int.Parse(Console.ReadLine());
                        Console.WriteLine("Ingrese el nombre de la base de datos en MongoDB (Ejem. Ventas)");
                        String dbName = Console.ReadLine();
                        mdlMetodos.ReadExcel(path, initTable, dbName);
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
                break;
            case 2:
                {
                    flag = false;
                }
                break;
            default:
                {
                    Console.WriteLine("Porfavor ingrese un valor valido");
                }
                break;
        }
    }catch(Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
    Console.WriteLine("Presione una tecla para continuar");
    Console.ReadKey();
    Console.Clear();
}