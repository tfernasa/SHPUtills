using System;
using SHPOperations;

namespace SHP
{
    class Program
    {
        const string site = "https://sharedshp.sharepoint.com/sites/develop";

        static void Main(string[] args)
        {
            Console.WriteLine("01 - Aplicación de TEST SHP.");
            Console.WriteLine("");

            Console.WriteLine("02 - Conectando SHP");

            using (SHPExecOperations connSHP = new SHPExecOperations(site, @"clopesam@sharedshp.onmicrosoft.com", "Pl4yb@ck", true))
            {
                if (connSHP.CreateList("TestTFS", 100, "", true))
                {
                    Console.WriteLine(string.Format("03 - Nueva tabla creada correctamente: {0}", connSHP.ListName));
                    Console.WriteLine("");

                    //Creando campos...
                    //UNA LÍNEA DE TEXTO
                    connSHP.AddNewColumn("Columna de prueba 1", false, false, 255);
                    connSHP.AddNewColumn("Columna de prueba 2", true, true, 50, "Hola mundo!", connSHP.ListName, "Descripción de la columna 2");

                    //VARIAS LÍNEAS DE TEXTO
                    connSHP.AddNewColumn("Columna de prueba 3", false, 0, false);
                    connSHP.AddNewColumn("Columna de prueba 4", true, 12, true, connSHP.ListName, "Descripción de la columna 4");

                    //ELECCION
                    string[] values = { "White", "Black", "Grey", "Blue", "Red", "Green", "Yellow" };
                    connSHP.AddNewColumn("Columna de prueba 5", false, false, values);
                    connSHP.AddNewColumn("Columna de prueba 6", true, true, values, "Black");

                    if (connSHP.CreateColumns())
                    {
                        Console.WriteLine(string.Format("04 - Campos añadidos correctamente a la lista {0}", connSHP.ListName));
                        Console.WriteLine("");
                    }
                }

            }

            Console.WriteLine("Presione cualquier tecla para finalizar");
            Console.ReadKey();
        }
    }
}
